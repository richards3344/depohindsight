import os
import uuid
import io
import threading
import time
import json
from datetime import datetime
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_file, Response, stream_with_context
)
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from flask_login import (
    LoginManager, UserMixin, login_user, logout_user,
    login_required, current_user
)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-change-me')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///depohindsight.db')
if app.config['SQLALCHEMY_DATABASE_URI'].startswith('postgres://'):
    app.config['SQLALCHEMY_DATABASE_URI'] = app.config['SQLALCHEMY_DATABASE_URI'].replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ── Models ──

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(255), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    jobs = db.relationship('Job', backref='user', lazy=True)


class Job(db.Model):
    id = db.Column(db.String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    filename = db.Column(db.String(255), nullable=False)
    deponent = db.Column(db.String(255))
    job_type = db.Column(db.String(20), nullable=False)
    status = db.Column(db.String(20), default='processing')
    progress = db.Column(db.Integer, default=0)
    progress_text = db.Column(db.String(255), default='Starting...')
    error_message = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    completed_at = db.Column(db.DateTime)
    documents = db.relationship('Document', backref='job', lazy=True)


class Document(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    job_id = db.Column(db.String(36), db.ForeignKey('job.id'), nullable=False)
    doc_type = db.Column(db.String(20), nullable=False)
    filename = db.Column(db.String(255), nullable=False)
    file_data = db.Column(db.LargeBinary, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


# ── CSRF ──

@app.before_request
def csrf_protect():
    if request.method == 'POST':
        token = request.form.get('csrf_token')
        if not token or token != request.cookies.get('csrf_token'):
            flash('Session expired. Please try again.', 'error')
            return redirect(request.url)


@app.after_request
def set_csrf_cookie(response):
    if 'csrf_token' not in request.cookies:
        token = uuid.uuid4().hex
        response.set_cookie('csrf_token', token, httponly=True, samesite='Lax')
    return response


@app.context_processor
def inject_csrf():
    token = request.cookies.get('csrf_token', uuid.uuid4().hex)
    return {'csrf_token': token}


# ── Routes ──

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        user = User.query.filter_by(email=email).first()
        if user and bcrypt.check_password_hash(user.password_hash, password):
            login_user(user)
            return redirect(url_for('dashboard'))
        flash('Invalid email or password.', 'error')
    return render_template('login.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        password2 = request.form.get('password2', '')
        if not email or not password:
            flash('All fields are required.', 'error')
        elif password != password2:
            flash('Passwords do not match.', 'error')
        elif len(password) < 6:
            flash('Password must be at least 6 characters.', 'error')
        elif User.query.filter_by(email=email).first():
            flash('An account with that email already exists.', 'error')
        else:
            pw_hash = bcrypt.generate_password_hash(password).decode('utf-8')
            user = User(email=email, password_hash=pw_hash)
            db.session.add(user)
            db.session.commit()
            flash('Account created. Please sign in.', 'success')
            return redirect(url_for('login'))
    return render_template('register.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('index'))


@app.route('/dashboard')
@login_required
def dashboard():
    jobs = Job.query.filter_by(user_id=current_user.id).order_by(Job.created_at.desc()).all()
    return render_template('dashboard.html', jobs=jobs)


@app.route('/upload', methods=['POST'])
@login_required
def upload():
    file = request.files.get('transcript')
    job_type = request.form.get('job_type', 'both')

    if not file or not file.filename:
        flash('Please select a file.', 'error')
        return redirect(url_for('dashboard'))

    if not file.filename.lower().endswith('.txt'):
        flash('Only .txt files are supported.', 'error')
        return redirect(url_for('dashboard'))

    if job_type not in ('summary', 'hindsight', 'both'):
        job_type = 'both'

    transcript_text = file.read().decode('utf-8', errors='replace')
    original_name = file.filename

    job = Job(
        user_id=current_user.id,
        filename=original_name,
        job_type=job_type,
    )
    db.session.add(job)
    db.session.commit()

    job_id = job.id
    thread = threading.Thread(target=run_analysis, args=(job_id, transcript_text, job_type), daemon=True)
    thread.start()

    return redirect(url_for('processing', job_id=job_id))


@app.route('/processing/<job_id>')
@login_required
def processing(job_id):
    job = db.session.get(Job, job_id)
    if not job or job.user_id != current_user.id:
        flash('Job not found.', 'error')
        return redirect(url_for('dashboard'))
    if job.status == 'complete':
        return redirect(url_for('dashboard'))
    return render_template('processing.html', job=job)


@app.route('/stream/<job_id>')
@login_required
def stream(job_id):
    job = db.session.get(Job, job_id)
    if not job or job.user_id != current_user.id:
        return Response('Unauthorized', status=403)

    def generate():
        while True:
            db.session.expire_all()
            j = db.session.get(Job, job_id)
            if not j:
                break

            data = {
                'progress': j.progress,
                'text': j.progress_text or '',
                'status': j.status,
            }

            if j.status == 'complete':
                docs = []
                for d in j.documents:
                    docs.append({'id': d.id, 'type': d.doc_type})
                data['documents'] = docs
                yield f"data: {json.dumps(data)}\n\n"
                break

            if j.status == 'failed':
                data['error'] = j.error_message or 'Analysis failed'
                yield f"data: {json.dumps(data)}\n\n"
                break

            yield f"data: {json.dumps(data)}\n\n"
            time.sleep(1.5)

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no',
        }
    )


@app.route('/download/<int:doc_id>')
@login_required
def download(doc_id):
    doc = db.session.get(Document, doc_id)
    if not doc:
        flash('Document not found.', 'error')
        return redirect(url_for('dashboard'))
    job = db.session.get(Job, doc.job_id)
    if not job or job.user_id != current_user.id:
        flash('Unauthorized.', 'error')
        return redirect(url_for('dashboard'))
    return send_file(
        io.BytesIO(doc.file_data),
        as_attachment=True,
        download_name=doc.filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


# ── Background processing ──

def run_analysis(job_id, transcript_text, job_type):
    from depo_engine import (
        parse_transcript, extract_name_from_transcript, extract_case_caption,
        generate_summary_claude, generate_critique_claude,
        create_document, create_critique_document,
    )

    with app.app_context():
        job = db.session.get(Job, job_id)
        if not job:
            return

        def update_progress(pct, text):
            job.progress = pct
            job.progress_text = text
            db.session.commit()

        def log(msg):
            job.progress_text = msg
            db.session.commit()

        try:
            update_progress(2, 'Parsing transcript...')
            qa_pairs = parse_transcript(transcript_text)
            deponent = extract_name_from_transcript(transcript_text)
            case_caption = extract_case_caption(transcript_text)
            job.deponent = deponent
            db.session.commit()

            update_progress(5, 'Transcript parsed. Starting AI analysis...')

            if job_type in ('summary', 'both'):
                config = {'admissions': 15, 'sections': 8, 'topics_per': 5}

                def summary_progress(pct, text):
                    if job_type == 'both':
                        scaled = 5 + int(pct * 0.4)
                    else:
                        scaled = 5 + int(pct * 0.9)
                    update_progress(scaled, text)

                summary_result = generate_summary_claude(
                    transcript_text, deponent, qa_pairs, log, config,
                    progress_func=summary_progress
                )
                if summary_result:
                    doc = create_document(summary_result, deponent, "Page:Line", qa_pairs, case_caption)
                    buf = io.BytesIO()
                    doc.save(buf)
                    fname = f"{deponent.replace(' ', '_')}_Summary.docx"
                    db_doc = Document(
                        job_id=job_id,
                        doc_type='Summary',
                        filename=fname,
                        file_data=buf.getvalue(),
                    )
                    db.session.add(db_doc)
                    db.session.commit()
                else:
                    if job_type == 'summary':
                        raise Exception('Summary generation failed — AI did not return valid data')

            if job_type in ('hindsight', 'both'):
                def hindsight_progress(pct, text):
                    if job_type == 'both':
                        scaled = 50 + int(pct * 0.45)
                    else:
                        scaled = 5 + int(pct * 0.9)
                    update_progress(scaled, text)

                critique_result = generate_critique_claude(
                    transcript_text, deponent, log,
                    progress_func=hindsight_progress
                )
                if critique_result:
                    doc = create_critique_document(critique_result, deponent, case_caption)
                    buf = io.BytesIO()
                    doc.save(buf)
                    fname = f"{deponent.replace(' ', '_')}_Hindsight.docx"
                    db_doc = Document(
                        job_id=job_id,
                        doc_type='Hindsight',
                        filename=fname,
                        file_data=buf.getvalue(),
                    )
                    db.session.add(db_doc)
                    db.session.commit()
                else:
                    if job_type == 'hindsight':
                        raise Exception('Hindsight generation failed — AI did not return valid data')

            job.status = 'complete'
            job.progress = 100
            job.progress_text = 'Complete!'
            job.completed_at = datetime.utcnow()
            db.session.commit()

        except Exception as e:
            job.status = 'failed'
            job.error_message = str(e)
            job.progress_text = 'Failed'
            db.session.commit()


# ── Init ──

with app.app_context():
    db.create_all()


if __name__ == '__main__':
    app.run(debug=True, port=8080)

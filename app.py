import os, random, secrets, io
from datetime import datetime, timezone
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook, Workbook

app=Flask(__name__)
app.secret_key=os.environ.get('SECRET_KEY') or secrets.token_hex(32)
app.config['MAX_CONTENT_LENGTH']=5*1024*1024
url=os.environ.get('DATABASE_URL','sqlite:///vacs_security.db')
if url.startswith('postgres://'): url='postgresql+psycopg://'+url[len('postgres://'):]
if url.startswith('postgresql://'): url='postgresql+psycopg://'+url[len('postgresql://'):]
app.config['SQLALCHEMY_DATABASE_URI']=url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS']=False
db=SQLAlchemy(app)

class Admin(db.Model):
    id=db.Column(db.Integer,primary_key=True); username=db.Column(db.String(80),unique=True,nullable=False); password_hash=db.Column(db.String(255),nullable=False)
class Question(db.Model):
    id=db.Column(db.Integer,primary_key=True); text=db.Column(db.Text,nullable=False); a=db.Column(db.Text,nullable=False); b=db.Column(db.Text,nullable=False); c=db.Column(db.Text,nullable=False); d=db.Column(db.Text,nullable=False); answer=db.Column(db.String(1),nullable=False); category=db.Column(db.String(120),default='Chung')
class Exam(db.Model):
    id=db.Column(db.Integer,primary_key=True); title=db.Column(db.String(200),nullable=False); question_count=db.Column(db.Integer,default=20); duration=db.Column(db.Integer,default=20); active=db.Column(db.Boolean,default=True); created_at=db.Column(db.DateTime,default=lambda:datetime.now(timezone.utc))
class Result(db.Model):
    id=db.Column(db.Integer,primary_key=True); candidate=db.Column(db.String(150),nullable=False); employee_id=db.Column(db.String(80)); exam_title=db.Column(db.String(200)); score=db.Column(db.Float); correct=db.Column(db.Integer); total=db.Column(db.Integer); submitted_at=db.Column(db.DateTime,default=lambda:datetime.now(timezone.utc))

def admin_required(f):
    @wraps(f)
    def w(*args,**kwargs):
        if not session.get('admin'): return redirect(url_for('login'))
        return f(*args,**kwargs)
    return w

@app.route('/')
def index(): return render_template('index.html', exams=Exam.query.filter_by(active=True).all())
@app.route('/login',methods=['GET','POST'])
def login():
    if request.method=='POST':
        a=Admin.query.filter_by(username=request.form['username']).first()
        if a and check_password_hash(a.password_hash,request.form['password']): session['admin']=a.username; return redirect(url_for('admin'))
        flash('Sai tài khoản hoặc mật khẩu.','danger')
    return render_template('login.html')
@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('index'))
@app.route('/admin')
@admin_required
def admin(): return render_template('admin.html', qcount=Question.query.count(), exams=Exam.query.all(), results=Result.query.order_by(Result.id.desc()).limit(10).all())
@app.route('/admin/exam',methods=['POST'])
@admin_required
def create_exam():
    e=Exam(title=request.form['title'],question_count=max(1,int(request.form['count'])),duration=max(1,int(request.form['duration'])))
    db.session.add(e); db.session.commit(); flash('Đã tạo đề thi.','success'); return redirect(url_for('admin'))
@app.route('/admin/upload',methods=['POST'])
@admin_required
def upload_questions():
    f=request.files.get('file')
    if not f: flash('Chưa chọn file.','danger'); return redirect(url_for('admin'))
    try:
        wb=load_workbook(f,data_only=True); ws=wb.active; added=0
        for row in ws.iter_rows(min_row=2,values_only=True):
            if not row or not row[0]: continue
            vals=list(row)+[None]*7
            ans=str(vals[5] or '').strip().upper()
            if ans not in ['A','B','C','D']: continue
            db.session.add(Question(text=str(vals[0]),a=str(vals[1]),b=str(vals[2]),c=str(vals[3]),d=str(vals[4]),answer=ans,category=str(vals[6] or 'Chung'))); added+=1
        db.session.commit(); flash(f'Đã nhập {added} câu hỏi.','success')
    except Exception as e: db.session.rollback(); flash('Không đọc được file Excel. Kiểm tra đúng mẫu.','danger')
    return redirect(url_for('admin'))
@app.route('/exam/<int:exam_id>',methods=['GET','POST'])
def take_exam(exam_id):
    exam=Exam.query.get_or_404(exam_id)
    if request.method=='POST':
        candidate=request.form.get('candidate','').strip(); employee_id=request.form.get('employee_id','').strip()
        if not candidate: flash('Vui lòng nhập họ tên.','danger'); return redirect(url_for('take_exam',exam_id=exam_id))
        qs=Question.query.all()
        if not qs: flash('Ngân hàng câu hỏi chưa có dữ liệu.','danger'); return redirect(url_for('index'))
        selected=random.sample(qs,min(exam.question_count,len(qs))); session['quiz']=[q.id for q in selected]; session['candidate']=candidate; session['employee_id']=employee_id; session['exam_id']=exam.id
        return render_template('quiz.html',exam=exam,questions=selected)
    return render_template('start_exam.html',exam=exam)
@app.route('/submit',methods=['POST'])
def submit():
    ids=session.get('quiz',[]); exam=Exam.query.get(session.get('exam_id'))
    if not ids or not exam: return redirect(url_for('index'))
    qs=Question.query.filter(Question.id.in_(ids)).all(); correct=sum(1 for q in qs if request.form.get(f'q{q.id}')==q.answer); total=len(qs); score=round(correct*10/total,2) if total else 0
    r=Result(candidate=session.get('candidate',''),employee_id=session.get('employee_id',''),exam_title=exam.title,score=score,correct=correct,total=total); db.session.add(r); db.session.commit(); session.pop('quiz',None)
    return render_template('result.html',result=r)
@app.route('/admin/results.xlsx')
@admin_required
def export_results():
    wb=Workbook(); ws=wb.active; ws.title='Ket qua'; ws.append(['STT','Họ tên','Mã NV','Đề thi','Điểm','Đúng','Tổng câu','Thời gian'])
    for i,r in enumerate(Result.query.order_by(Result.id).all(),1): ws.append([i,r.candidate,r.employee_id,r.exam_title,r.score,r.correct,r.total,r.submitted_at.strftime('%d/%m/%Y %H:%M') if r.submitted_at else ''])
    bio=io.BytesIO(); wb.save(bio); bio.seek(0); return send_file(bio,as_attachment=True,download_name='ket_qua_VACS_SECURITY.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

with app.app_context():
    db.create_all()
    if not Admin.query.first():
        db.session.add(Admin(username=os.environ.get('ADMIN_USER','admin'), password_hash=generate_password_hash(os.environ.get('ADMIN_PASSWORD','Vacs@123'))))
        db.session.commit()

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)),debug=True)

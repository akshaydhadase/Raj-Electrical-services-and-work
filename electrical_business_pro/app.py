import os
from datetime import datetime
from pathlib import Path
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

BASE=Path(__file__).resolve().parent
(BASE/'instance').mkdir(exist_ok=True); (BASE/'static/uploads').mkdir(parents=True,exist_ok=True)
app=Flask(__name__); app.config.update(SECRET_KEY=os.getenv('SECRET_KEY','change-me'),SQLALCHEMY_DATABASE_URI='sqlite:///'+str(BASE/'instance/electrical.db'),SQLALCHEMY_TRACK_MODIFICATIONS=False,MAX_CONTENT_LENGTH=50*1024*1024)
db=SQLAlchemy(app)

class Admin(db.Model):
 id=db.Column(db.Integer,primary_key=True); username=db.Column(db.String(80),unique=True,nullable=False); password_hash=db.Column(db.String(255),nullable=False)
class Setting(db.Model):
 id=db.Column(db.Integer,primary_key=True); company_name=db.Column(db.String(150),default='Spark Electrical Solutions'); tagline=db.Column(db.String(250),default='Professional Electrical & Tender Services'); phone=db.Column(db.String(50),default='+91 9860991315'); email=db.Column(db.String(150),default='info@example.com'); instagram=db.Column(db.String(250),default='https://instagram.com/'); whatsapp=db.Column(db.String(50),default=''); address=db.Column(db.Text,default='Your business address'); about=db.Column(db.Text,default='Reliable electrical contracting, maintenance and project services.'); hero_title=db.Column(db.String(250),default='Powering Projects. Building Trust.'); hero_text=db.Column(db.Text,default='Electrical contracting, maintenance, industrial solutions and tender support.')
class Service(db.Model):
 id=db.Column(db.Integer,primary_key=True); title=db.Column(db.String(150),nullable=False); description=db.Column(db.Text,default=''); icon=db.Column(db.String(20),default='⚡'); active=db.Column(db.Boolean,default=True)
class Product(db.Model):
 id=db.Column(db.Integer,primary_key=True); code=db.Column(db.String(80),default=''); name=db.Column(db.String(150),nullable=False); brand=db.Column(db.String(100),default=''); price=db.Column(db.Float,default=0); stock=db.Column(db.Integer,default=0); min_stock=db.Column(db.Integer,default=5); active=db.Column(db.Boolean,default=True)
class Project(db.Model):
 id=db.Column(db.Integer,primary_key=True); title=db.Column(db.String(180),nullable=False); client=db.Column(db.String(150),default=''); location=db.Column(db.String(150),default=''); project_type=db.Column(db.String(120),default=''); value=db.Column(db.Float,default=0); completion_date=db.Column(db.Date,nullable=True); image=db.Column(db.String(255),default=''); description=db.Column(db.Text,default=''); active=db.Column(db.Boolean,default=True)
class Tender(db.Model):
 id=db.Column(db.Integer,primary_key=True); tender_no=db.Column(db.String(100),nullable=False); client=db.Column(db.String(150),nullable=False); department=db.Column(db.String(150),default=''); title=db.Column(db.String(200),default=''); opening_date=db.Column(db.Date,nullable=True); closing_date=db.Column(db.Date,nullable=True); estimated_value=db.Column(db.Float,default=0); quoted_value=db.Column(db.Float,default=0); status=db.Column(db.String(30),default='New'); remarks=db.Column(db.Text,default='')
class Customer(db.Model):
 id=db.Column(db.Integer,primary_key=True); name=db.Column(db.String(150),nullable=False); phone=db.Column(db.String(50),default=''); email=db.Column(db.String(150),default=''); company=db.Column(db.String(150),default=''); status=db.Column(db.String(40),default='Lead'); created_at=db.Column(db.DateTime,default=datetime.utcnow)
class Enquiry(db.Model):
 id=db.Column(db.Integer,primary_key=True); name=db.Column(db.String(150),nullable=False); phone=db.Column(db.String(50),default=''); email=db.Column(db.String(150),default=''); service=db.Column(db.String(150),default=''); message=db.Column(db.Text,default=''); status=db.Column(db.String(40),default='New'); created_at=db.Column(db.DateTime,default=datetime.utcnow)
class Review(db.Model):
 id=db.Column(db.Integer,primary_key=True); customer_name=db.Column(db.String(150),nullable=False); rating=db.Column(db.Integer,default=5); comment=db.Column(db.Text,default=''); approved=db.Column(db.Boolean,default=False); created_at=db.Column(db.DateTime,default=datetime.utcnow)
class CompanyMedia(db.Model):
 id=db.Column(db.Integer,primary_key=True); title=db.Column(db.String(180),default=''); filename=db.Column(db.String(255),nullable=False); media_type=db.Column(db.String(20),nullable=False); active=db.Column(db.Boolean,default=True); created_at=db.Column(db.DateTime,default=datetime.utcnow)
class Quotation(db.Model):
 id=db.Column(db.Integer,primary_key=True); quote_no=db.Column(db.String(100),nullable=False); customer=db.Column(db.String(150),nullable=False); phone=db.Column(db.String(50),default=''); subtotal=db.Column(db.Float,default=0); discount=db.Column(db.Float,default=0); gst=db.Column(db.Float,default=0); total=db.Column(db.Float,default=0); status=db.Column(db.String(40),default='Draft'); created_at=db.Column(db.DateTime,default=datetime.utcnow)

def site():
 s=Setting.query.first()
 if not s: s=Setting(); db.session.add(s); db.session.commit()
 return s
@app.context_processor
def globals(): return {'site':site()}
def auth(f):
 @wraps(f)
 def w(*a,**k): return f(*a,**k) if session.get('admin_id') else redirect(url_for('login'))
 return w
def dt(v):
 try:return datetime.strptime(v,'%Y-%m-%d').date() if v else None
 except:return None

def seed():
 db.create_all()
 if not Admin.query.first(): db.session.add(Admin(username='ujwaladhadase',password_hash=generate_password_hash('Umesh@dh#99')))
 site()
 if not Service.query.first(): db.session.add_all([Service(title='Industrial Electrical Work',description='Panels, wiring, installation and maintenance.',icon='🏭'),Service(title='Commercial Electrical',description='Office, shop and commercial electrical solutions.',icon='🏢'),Service(title='Tender & Project Support',description='Tender tracking, quotations and project support.',icon='📋'),Service(title='Electrical Maintenance',description='Preventive maintenance and troubleshooting.',icon='🛠️')])
 if not Product.query.first(): db.session.add_all([Product(code='CB-001',name='MCB Circuit Breaker',brand='Demo Brand',price=450,stock=50,min_stock=10),Product(code='CAB-001',name='Copper Cable',brand='Demo Brand',price=1200,stock=25,min_stock=5)])
 db.session.commit()

@app.route('/')
def home(): return render_template('customer/home.html',services=Service.query.filter_by(active=True).all(),projects=Project.query.filter_by(active=True).order_by(Project.id.desc()).limit(6).all(),products=Product.query.filter_by(active=True).limit(8).all(),reviews=Review.query.filter_by(approved=True).order_by(Review.id.desc()).limit(6).all(),media=CompanyMedia.query.filter_by(active=True).order_by(CompanyMedia.id.desc()).all())
@app.route('/services')
def services(): return render_template('customer/services.html',services=Service.query.filter_by(active=True).all())
@app.route('/products')
def products(): return render_template('customer/products.html',products=Product.query.filter_by(active=True).all())
@app.route('/projects')
def projects(): return render_template('customer/projects.html',projects=Project.query.filter_by(active=True).all())
@app.route('/contact',methods=['GET','POST'])
def contact():
 if request.method=='POST':
  e=Enquiry(name=request.form.get('name','').strip(),phone=request.form.get('phone','').strip(),email=request.form.get('email','').strip(),service=request.form.get('service','').strip(),message=request.form.get('message','').strip())
  if not e.name: flash('Name is required.','danger')
  else:
   db.session.add(e); existing=Customer.query.filter_by(phone=e.phone).first() if e.phone else None
   if not existing: db.session.add(Customer(name=e.name,phone=e.phone,email=e.email))
   db.session.commit(); flash('Enquiry submitted successfully.','success'); return redirect(url_for('contact'))
 return render_template('customer/contact.html',services=Service.query.filter_by(active=True).all())
@app.route('/review',methods=['POST'])
def review():
 r=Review(customer_name=request.form.get('customer_name','').strip(),rating=max(1,min(5,int(request.form.get('rating',5)))),comment=request.form.get('comment','').strip()); db.session.add(r); db.session.commit(); flash('Review sent for approval.','success'); return redirect(url_for('home'))

@app.route('/admin/login',methods=['GET','POST'])
def login():
 if request.method=='POST':
  a=Admin.query.filter_by(username=request.form.get('username','').strip()).first()
  if a and check_password_hash(a.password_hash,request.form.get('password','')): session['admin_id']=a.id; return redirect(url_for('dashboard'))
  flash('Invalid login.','danger')
 return render_template('admin/login.html')
@app.route('/admin/logout')
def logout(): session.clear(); return redirect(url_for('login'))
@app.route('/admin')
@auth
def dashboard():
 ts=Tender.query.all(); won=sum(t.status=='Won' for t in ts); value=sum(t.quoted_value or 0 for t in ts)
 return render_template('admin/dashboard.html',customers=Customer.query.count(),enquiries=Enquiry.query.count(),tenders=len(ts),won=won,win_rate=round(won*100/len(ts),1) if ts else 0,tender_value=value,low_stock=Product.query.filter(Product.stock<=Product.min_stock).count(),recent=Enquiry.query.order_by(Enquiry.id.desc()).limit(8).all())
@app.route('/admin/settings',methods=['GET','POST'])
@auth
def settings():
 s=site()
 if request.method=='POST':
  for x in ['company_name','tagline','phone','email','instagram','whatsapp','address','about','hero_title','hero_text']:
   setattr(s,x,request.form.get(x,''))
  db.session.commit()
  flash('Company profile and contact information updated successfully.','success')
  return redirect(url_for('settings'))
 return render_template('admin/settings.html',s=s)
@app.route('/admin/password',methods=['GET','POST'])
@auth
def password():
 a=db.session.get(Admin,session['admin_id'])
 if request.method=='POST':
  if not check_password_hash(a.password_hash,request.form.get('current','')): flash('Current password incorrect.','danger')
  elif len(request.form.get('new',''))<8: flash('Password must be 8+ characters.','danger')
  else: a.password_hash=generate_password_hash(request.form['new']); db.session.commit(); flash('Password changed.','success')
 return render_template('admin/password.html')

def simple_crud(model, template, fields, title):
 pass
@app.route('/admin/services',methods=['GET','POST'])
@auth
def admin_services():
 if request.method=='POST': db.session.add(Service(title=request.form['title'],description=request.form.get('description',''),icon=request.form.get('icon','⚡'),active=bool(request.form.get('active')))); db.session.commit(); return redirect(url_for('admin_services'))
 return render_template('admin/services.html',items=Service.query.order_by(Service.id.desc()).all())
@app.post('/admin/services/delete/<int:id>')
@auth
def service_delete(id):
 x=db.session.get(Service,id)
 if x: db.session.delete(x); db.session.commit()
 return redirect(url_for('admin_services'))
@app.route('/admin/products',methods=['GET','POST'])
@auth
def admin_products():
 if request.method=='POST': db.session.add(Product(code=request.form.get('code',''),name=request.form['name'],brand=request.form.get('brand',''),price=float(request.form.get('price',0) or 0),stock=int(request.form.get('stock',0) or 0),min_stock=int(request.form.get('min_stock',5) or 5),active=bool(request.form.get('active')))); db.session.commit(); return redirect(url_for('admin_products'))
 return render_template('admin/products.html',items=Product.query.order_by(Product.id.desc()).all())
@app.post('/admin/products/delete/<int:id>')
@auth
def product_delete(id):
 x=db.session.get(Product,id)
 if x: db.session.delete(x); db.session.commit()
 return redirect(url_for('admin_products'))

MEDIA_IMAGE_EXT={'jpg','jpeg','png','webp','gif'}
MEDIA_VIDEO_EXT={'mp4','webm','mov','avi','mkv'}

def save_media_file(f, allowed):
 if not f or not f.filename: return None
 ext=f.filename.lower().rsplit('.',1)[-1] if '.' in f.filename else ''
 if ext not in allowed: return None
 name=datetime.utcnow().strftime('%Y%m%d%H%M%S%f')+'_'+secure_filename(f.filename)
 f.save(BASE/'static/uploads'/name)
 return name

@app.route('/admin/media',methods=['GET','POST'])
@auth
def admin_media():
 if request.method=='POST':
  f=request.files.get('file')
  ext=f.filename.lower().rsplit('.',1)[-1] if f and '.' in f.filename else ''
  if ext in MEDIA_IMAGE_EXT: media_type='image'
  elif ext in MEDIA_VIDEO_EXT: media_type='video'
  else:
   flash('Please upload JPG, PNG, WEBP, GIF, MP4, WEBM, MOV, AVI or MKV.','danger')
   return redirect(url_for('admin_media'))
  name=save_media_file(f, MEDIA_IMAGE_EXT|MEDIA_VIDEO_EXT)
  db.session.add(CompanyMedia(title=request.form.get('title','').strip(),filename=name,media_type=media_type,active=bool(request.form.get('active'))))
  db.session.commit(); flash('Image/video uploaded successfully.','success'); return redirect(url_for('admin_media'))
 return render_template('admin/media.html',items=CompanyMedia.query.order_by(CompanyMedia.id.desc()).all())

@app.post('/admin/media/delete/<int:id>')
@auth
def media_delete(id):
 x=db.session.get(CompanyMedia,id)
 if x:
  fp=BASE/'static/uploads'/x.filename
  if fp.exists(): fp.unlink()
  db.session.delete(x); db.session.commit()
 return redirect(url_for('admin_media'))

@app.route('/admin/projects',methods=['GET','POST'])
@auth
def admin_projects():
 if request.method=='POST':
  p=Project(title=request.form['title'],client=request.form.get('client',''),location=request.form.get('location',''),project_type=request.form.get('project_type',''),value=float(request.form.get('value',0) or 0),completion_date=dt(request.form.get('completion_date')),description=request.form.get('description',''),active=bool(request.form.get('active')))
  f=request.files.get('image')
  if f and f.filename and f.filename.lower().rsplit('.',1)[-1] in {'jpg','jpeg','png','webp','gif'}:
   name=datetime.utcnow().strftime('%Y%m%d%H%M%S%f')+'_'+secure_filename(f.filename); f.save(BASE/'static/uploads'/name); p.image=name
  db.session.add(p); db.session.commit(); return redirect(url_for('admin_projects'))
 return render_template('admin/projects.html',items=Project.query.order_by(Project.id.desc()).all())
@app.post('/admin/projects/delete/<int:id>')
@auth
def project_delete(id):
 x=db.session.get(Project,id)
 if x: db.session.delete(x); db.session.commit()
 return redirect(url_for('admin_projects'))
@app.route('/admin/tenders',methods=['GET','POST'])
@auth
def admin_tenders():
 if request.method=='POST': db.session.add(Tender(tender_no=request.form['tender_no'],client=request.form['client'],department=request.form.get('department',''),title=request.form.get('title',''),opening_date=dt(request.form.get('opening_date')),closing_date=dt(request.form.get('closing_date')),estimated_value=float(request.form.get('estimated_value',0) or 0),quoted_value=float(request.form.get('quoted_value',0) or 0),status=request.form.get('status','New'),remarks=request.form.get('remarks',''))); db.session.commit(); return redirect(url_for('admin_tenders'))
 return render_template('admin/tenders.html',items=Tender.query.order_by(Tender.id.desc()).all())
@app.post('/admin/tenders/delete/<int:id>')
@auth
def tender_delete(id):
 x=db.session.get(Tender,id)
 if x: db.session.delete(x); db.session.commit()
 return redirect(url_for('admin_tenders'))
@app.route('/admin/customers')
@auth
def customers(): return render_template('admin/customers.html',items=Customer.query.order_by(Customer.id.desc()).all())
@app.route('/admin/enquiries')
@auth
def enquiries(): return render_template('admin/enquiries.html',items=Enquiry.query.order_by(Enquiry.id.desc()).all())
@app.post('/admin/enquiries/status/<int:id>')
@auth
def enquiry_status(id):
 x=db.session.get(Enquiry,id)
 if x: x.status=request.form['status']; db.session.commit()
 return redirect(url_for('enquiries'))
@app.route('/admin/reviews')
@auth
def reviews(): return render_template('admin/reviews.html',items=Review.query.order_by(Review.id.desc()).all())
@app.post('/admin/reviews/approve/<int:id>')
@auth
def approve(id):
 x=db.session.get(Review,id)
 if x: x.approved=True; db.session.commit()
 return redirect(url_for('reviews'))
@app.post('/admin/reviews/delete/<int:id>')
@auth
def review_delete(id):
 x=db.session.get(Review,id)
 if x: db.session.delete(x); db.session.commit()
 return redirect(url_for('reviews'))
@app.route('/admin/quotations',methods=['GET','POST'])
@auth
def quotations():
 if request.method=='POST':
  q=Quotation(quote_no=request.form['quote_no'],customer=request.form['customer'],phone=request.form.get('phone',''),subtotal=float(request.form.get('subtotal',0) or 0),discount=float(request.form.get('discount',0) or 0),gst=float(request.form.get('gst',18) or 0),status=request.form.get('status','Draft')); q.total=max(0,q.subtotal-q.discount); q.total+=q.total*q.gst/100; db.session.add(q); db.session.commit(); return redirect(url_for('quotations'))
 return render_template('admin/quotations.html',items=Quotation.query.order_by(Quotation.id.desc()).all())

def excel(rows,headers,name):
 wb=Workbook(); ws=wb.active; ws.title='Report'; ws.append(headers)
 for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F4E78')
 for r in rows: ws.append(r)
 p=BASE/'instance'/name; wb.save(p); return send_file(p,as_attachment=True,download_name=name)
@app.route('/admin/reports/tenders.xlsx')
@auth
def tender_report():
 return excel([[t.tender_no,t.client,t.department,t.title,t.opening_date,t.closing_date,t.estimated_value,t.quoted_value,t.status,t.remarks] for t in Tender.query.all()],['Tender No','Client','Department','Title','Opening','Closing','Estimated','Quoted','Status','Remarks'],'Tender_Report.xlsx')
@app.route('/admin/reports/customers.xlsx')
@auth
def customer_report(): return excel([[c.name,c.phone,c.email,c.company,c.status,c.created_at] for c in Customer.query.all()],['Name','Phone','Email','Company','Status','Created'],'Customer_Report.xlsx')

if __name__=='__main__':
 with app.app_context(): seed()
 print('Customer: http://127.0.0.1:5000'); print('Admin: http://127.0.0.1:5000/admin/login'); #print('Login: admin / admin123')
 app.run(debug=True,host='127.0.0.1',port=5000)

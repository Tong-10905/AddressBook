import os
import io
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
from sqlalchemy import Column, Integer, String, Boolean, Text

# --- 初始化应用与数据库 ---
app = Flask(__name__)
app.secret_key = 'super_secret_key_for_flash_messages' # 用于闪现消息
# 使用 SQLite 数据库，数据会保存在当前目录的 addressbook.db 文件中
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///addressbook.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# --- 数据模型定义 ---
class Contact(db.Model):
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    is_bookmarked = Column(Boolean, default=False) # 书签功能
    # 为了方便，我们将多种联系方式存为一个文本格式，用逗号或分号分隔
    # 格式例如: "phone:13800000000; email:test@qq.com"
    contact_methods = Column(Text, default="") 

    def __repr__(self):
        return f'<Contact {self.name}>'

# 首次运行创建数据库表
with app.app_context():
    db.create_all()

# --- 路由：主页 (展示列表) ---
@app.route('/')
def index():
    # 获取所有联系人，书签的排在前面，然后按名字排序
    contacts = Contact.query.order_by(Contact.is_bookmarked.desc(), Contact.name).all()
    return render_template('index.html', contacts=contacts)

# --- 路由：添加联系人 ---
@app.route('/add', methods=['POST'])
def add_contact():
    name = request.form.get('name')
    methods = request.form.get('methods', '') # 获取输入的联系方式
    
    if name:
        new_contact = Contact(name=name, contact_methods=methods)
        db.session.add(new_contact)
        db.session.commit()
        flash('联系人添加成功！', 'success')
    else:
        flash('姓名不能为空', 'danger')
        
    return redirect(url_for('index'))

# --- 路由：切换书签状态 ---
@app.route('/toggle_bookmark/<int:id>')
def toggle_bookmark(id):
    contact = Contact.query.get_or_404(id)
    contact.is_bookmarked = not contact.is_bookmarked
    db.session.commit()
    return redirect(url_for('index'))

# --- 路由：删除联系人 ---
@app.route('/delete/<int:id>')
def delete_contact(id):
    contact = Contact.query.get_or_404(id)
    db.session.delete(contact)
    db.session.commit()
    flash('联系人已删除', 'info')
    return redirect(url_for('index'))

# --- 路由：导出 Excel ---
@app.route('/export')
def export_excel():
    contacts = Contact.query.all()
    
    # 将数据转换为 Pandas 可以处理的字典列表
    data = []
    for c in contacts:
        data.append({
            'Name': c.name,
            'IsBookmarked': 'Yes' if c.is_bookmarked else 'No',
            'ContactMethods': c.contact_methods
        })
    
    df = pd.DataFrame(data)
    
    # 如果没有数据，给个默认的表头
    if df.empty:
        df = pd.DataFrame(columns=['Name', 'IsBookmarked', 'ContactMethods'])
        
    # 导出到内存中
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Contacts')
    output.seek(0)
    
    return send_file(output, download_name="contacts_export.xlsx", as_attachment=True)

# --- 路由：导入 Excel ---
@app.route('/import', methods=['POST'])
def import_excel():
    if 'file' not in request.files:
        flash('未找到文件', 'danger')
        return redirect(url_for('index'))
        
    file = request.files['file']
    if file.filename == '':
        flash('未选择文件', 'danger')
        return redirect(url_for('index'))
        
    if file and file.filename.endswith(('.xls', '.xlsx')):
        try:
            # 读取 Excel
            df = pd.read_excel(file)
            
            # 遍历行并写入数据库
            count = 0
            for index, row in df.iterrows():
                name = str(row.get('Name', '')) if pd.notna(row.get('Name')) else ''
                is_bookmarked = str(row.get('IsBookmarked', '')) == 'Yes' if pd.notna(row.get('IsBookmarked')) else False
                methods = str(row.get('ContactMethods', '')) if pd.notna(row.get('ContactMethods')) else ''
                
                if name and name != 'nan':
                    new_contact = Contact(
                        name=name, 
                        is_bookmarked=is_bookmarked, 
                        contact_methods=methods
                    )
                    db.session.add(new_contact)
                    count += 1
            
            db.session.commit()
            flash(f'成功导入 {count} 条联系人数据！', 'success')
        except Exception as e:
            db.session.rollback()
            flash(f'导入失败: {str(e)}', 'danger')
    else:
        flash('请上传 .xlsx 或 .xls 格式的文件', 'warning')
        
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
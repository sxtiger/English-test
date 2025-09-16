from flask import Flask, render_template, request, session, redirect, url_for, flash
from flask_session import Session
from werkzeug.utils import secure_filename
import pandas as pd
import random
import os
import md2excel
from datetime import datetime
import logging
import shutil
import ssl

# 初始化Flask应用
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', os.urandom(24).hex())
app.jinja_env.globals.update(min=min, max=max)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

# 应用配置
app.config.update(
    SESSION_TYPE='filesystem',
    SESSION_FILE_DIR='./flask_sessions',
    SESSION_COOKIE_NAME='lang_test',
    SESSION_COOKIE_SECURE=True,  # HTTPS环境下启用
    SESSION_COOKIE_HTTPONLY=True,
    PERMANENT_SESSION_LIFETIME=1800,
    MD_FILE='English Phrase.md',
    EXCEL_FILE='English Phrase.xlsx',
    ERROR_EXCEL='Error_Phrases.xlsx',
    HTTPS_PORT=int(os.environ.get('HTTPS_PORT', 51001)),
    HTTP_PORT=int(os.environ.get('HTTP_PORT', 51000)),
    UPLOAD_FOLDER='./uploads',
    MAX_CONTENT_LENGTH=16 * 1024 * 1024,
    ALLOWED_EXTENSIONS={'md'}
)
Session(app)

# 全局会话存储
user_sessions = {}

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def save_error_phrase(phrase):
    """保存错题到Excel"""
    try:
        error_data = {
            "英文短语": phrase.get('英文短语', ''),
            "中文翻译": phrase.get('中文翻译', ''),
            "学习日": phrase.get('学习日', ''),
            "书本名称": phrase.get('书本名称', ''),
            "错误时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        df = pd.DataFrame([error_data])
        if os.path.exists(app.config['ERROR_EXCEL']):
            df_existing = pd.read_excel(app.config['ERROR_EXCEL'])
            df = pd.concat([df_existing, df], ignore_index=True)
        
        df.to_excel(app.config['ERROR_EXCEL'], index=False)
        app.logger.info("错题保存成功")
    except Exception as e:
        app.logger.error(f"错题保存失败: {str(e)}")

def generate_options(correct, pool):
    """生成题目选项"""
    try:
        unique_pool = list({v for v in pool if v != correct})
        wrong_answers = random.sample(unique_pool, 3) if len(unique_pool) >=3 \
                       else (unique_pool * 3)[:3]
        
        options = [correct] + wrong_answers
        random.shuffle(options)
        return options
    except Exception as e:
        app.logger.error(f"选项生成失败: {str(e)}")
        return [correct, "选项生成错误1", "选项生成错误2", "选项生成错误3"]

def create_sample_md_file():
    """创建示例Markdown文件"""
    sample_content = """# 示例词汇书
## Day 1
- hello: 你好
- world: 世界
- good: 好的
- morning: 早晨
- afternoon: 下午

## Day 2  
- evening: 晚上
- night: 夜晚
- thank you: 谢谢你
- please: 请
- sorry: 对不起
"""
    try:
        with open(app.config['MD_FILE'], 'w', encoding='utf-8') as f:
            f.write(sample_content)
        app.logger.info("创建示例MD文件成功")
    except Exception as e:
        app.logger.error(f"创建示例MD文件失败: {str(e)}")

@app.before_request
def force_https():
    """强制HTTPS重定向"""
    if not request.is_secure and request.url.startswith('http://'):
        url = request.url.replace('http://', 'https://', 1)
        url = url.replace(':51000', ':51001')  # 重定向到HTTPS端口
        return redirect(url, code=301)

@app.route('/')
def index():
    """首页路由"""
    try:
        if not os.path.exists(app.config['MD_FILE']):
            create_sample_md_file()
        
        normal_total = md2excel.markdown_to_excel(
            app.config['MD_FILE'],
            app.config['EXCEL_FILE']
        )
        
        error_total = 0
        if os.path.exists(app.config['ERROR_EXCEL']):
            error_df = pd.read_excel(app.config['ERROR_EXCEL'])
            error_total = len(error_df)
        
        session.clear()
        app.logger.info(f"首页加载成功 正常题:{normal_total} 错题:{error_total}")
        return render_template('index.html',
                             normal_total=normal_total,
                             error_total=error_total,
                             default=50)
        
    except Exception as e:
        app.logger.error(f"首页加载失败: {str(e)}")
        return render_template('error.html', message="系统初始化失败")

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传路由"""
    try:
        app.logger.debug("收到文件上传请求")
        
        if 'file' not in request.files:
            flash('没有选择文件')
            return redirect(url_for('index'))
        
        file = request.files['file']
        
        if file.filename == '':
            flash('没有选择文件')
            return redirect(url_for('index'))
        
        if not allowed_file(file.filename):
            flash('只允许上传.md文件')
            return redirect(url_for('index'))
        
        if not os.path.exists(app.config['UPLOAD_FOLDER']):
            os.makedirs(app.config['UPLOAD_FOLDER'])
        
        filename = secure_filename(file.filename)
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(temp_path)
        
        if os.path.exists(app.config['MD_FILE']):
            backup_name = f"English Phrase_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
            shutil.copy2(app.config['MD_FILE'], backup_name)
            app.logger.info(f"已备份原文件为: {backup_name}")
        
        shutil.move(temp_path, app.config['MD_FILE'])
        
        try:
            phrase_count = md2excel.markdown_to_excel(app.config['MD_FILE'], app.config['EXCEL_FILE'])
            flash(f'文件上传成功！共导入 {phrase_count} 个短语')
            app.logger.info(f"文件上传成功，导入 {phrase_count} 个短语")
        except Exception as convert_error:
            backup_files = [f for f in os.listdir('.') if f.startswith('English Phrase_backup_')]
            if backup_files:
                latest_backup = max(backup_files)
                shutil.copy2(latest_backup, app.config['MD_FILE'])
                app.logger.warning(f"文件格式错误，已恢复备份: {latest_backup}")
            else:
                create_sample_md_file()
                
            flash(f'文件格式错误，上传失败: {str(convert_error)}')
            app.logger.error(f"文件转换失败: {str(convert_error)}")
        
        return redirect(url_for('index'))
        
    except Exception as e:
        app.logger.error(f"文件上传失败: {str(e)}")
        flash(f'文件上传失败: {str(e)}')
        return redirect(url_for('index'))

@app.route('/start', methods=['POST'])
def start_test():
    """开始测试路由"""
    try:
        app.logger.debug("收到开始测试请求")
        form_data = request.form
        app.logger.debug(f"表单数据: {dict(form_data)}")
        
        test_mode = form_data.get('test_mode', 'normal')
        question_count = int(form_data.get('question_count', 50))
        app.logger.info(f"测试模式: {test_mode}, 请求题数: {question_count}")

        if test_mode == 'error':
            if not os.path.exists(app.config['ERROR_EXCEL']):
                raise ValueError("错题库不存在，请先进行正常测试")
            
            df = pd.read_excel(app.config['ERROR_EXCEL'])
            if df.empty:
                raise ValueError("错题库为空")
                
            total = len(df)
            question_count = min(max(1, question_count), total)
            test_data = df.sample(n=question_count).to_dict('records')
        else:
            md2excel.markdown_to_excel(app.config['MD_FILE'], app.config['EXCEL_FILE'])
            df = pd.read_excel(app.config['EXCEL_FILE'])
            total = len(df)
            question_count = min(max(1, question_count), total)
            test_data = df.sample(n=question_count).to_dict('records')

        session.clear()
        session['test_mode'] = test_mode
        session['question_count'] = question_count
        session['user_id'] = os.urandom(16).hex()
        
        user_sessions[session['user_id']] = {
            'test_data': test_data,
            'current': 0,
            'correct': 0,
            'wrong': 0,
            'wrong_list': []
        }
        app.logger.info(f"会话初始化成功 用户ID: {session['user_id']}")
        return redirect(url_for('show_question'))

    except ValueError as ve:
        error_msg = str(ve)
        app.logger.warning(f"参数错误: {error_msg}")
        return render_template('error.html', message=error_msg)
    except Exception as e:
        error_msg = f"系统错误: {str(e)}"
        app.logger.error(error_msg)
        return render_template('error.html', message=error_msg)

@app.route('/test', methods=['GET', 'POST'])
def show_question():
    """题目展示路由"""
    try:
        user_id = session.get('user_id')
        if not user_id or user_id not in user_sessions:
            app.logger.warning("无效会话访问")
            return redirect(url_for('index'))
        
        user_data = user_sessions[user_id]
        app.logger.debug(f"当前进度: {user_data['current']}/{len(user_data['test_data'])}")

        if request.method == 'POST':
            current = user_data['current']
            question = user_data['test_data'][current]
            user_choice = request.form.get('answer', '0')
            
            if user_choice == str(question.get('correct_index', 0)):
                user_data['correct'] += 1
                app.logger.debug(f"第{current+1}题回答正确")
            else:
                app.logger.debug(f"第{current+1}题回答错误")
                if session.get('test_mode') == 'normal':
                    save_error_phrase(question)
                user_data['wrong'] += 1
                user_data['wrong_list'].append({
                    'question': question.get('display', ''),
                    'correct': question.get('correct_answer', ''),
                    'selected': question['options'][int(user_choice)-1] if user_choice.isdigit() else '无效选择',
                    'options': question.get('options', [])
                })
            user_data['current'] += 1

        if user_data['current'] >= len(user_data['test_data']):
            app.logger.info(f"测试完成 正确率: {user_data['correct']}/{len(user_data['test_data'])}")
            return redirect(url_for('show_result'))

        current = user_data['current']
        item = user_data['test_data'][current]
        is_english = random.choice([True, False])
        
        if is_english:
            question = item.get('英文短语', '')
            correct = item.get('中文翻译', '')
            pool = [p.get('中文翻译', '') for p in user_data['test_data']]
        else:
            question = item.get('中文翻译', '')
            correct = item.get('英文短语', '')
            pool = [p.get('英文短语', '') for p in user_data['test_data']]
        
        options = generate_options(correct, pool)
        item.update({
            'display': question,
            'options': options,
            'correct_index': options.index(correct) + 1,
            'correct_answer': correct
        })
        
        app.logger.debug(f"生成第{current+1}题: {question}")
        return render_template('test.html',
                             question=question,
                             options=options,
                             progress=current+1,
                             total=len(user_data['test_data']))

    except Exception as e:
        app.logger.error(f"题目展示失败: {str(e)}")
        return redirect(url_for('reset_test'))

@app.route('/result')
def show_result():
    """结果展示路由"""
    try:
        user_id = session.get('user_id')
        if not user_id or user_id not in user_sessions:
            return redirect(url_for('index'))
        
        user_data = user_sessions[user_id]
        return render_template('result.html',
                             correct=user_data['correct'],
                             wrong=user_data['wrong'],
                             wrong_list=user_data['wrong_list'],
                             total=len(user_data['test_data']))
    except Exception as e:
        app.logger.error(f"结果页加载失败: {str(e)}")
        return redirect(url_for('index'))

@app.route('/reset')
def reset_test():
    """重置测试路由"""
    user_id = session.get('user_id')
    if user_id and user_id in user_sessions:
        del user_sessions[user_id]
    session.clear()
    app.logger.info("测试已重置")
    return redirect(url_for('index'))

@app.route('/health')
def health_check():
    """健康检查路由"""
    return {'status': 'healthy', 'timestamp': datetime.now().isoformat()}

@app.after_request
def add_header(response):
    """添加安全头部"""
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    # HTTPS安全头部
    response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-XSS-Protection"] = "1; mode=block"
    return response

if __name__ == '__main__':
    # 初始化目录
    for directory in [app.config['SESSION_FILE_DIR'], app.config['UPLOAD_FOLDER'], './ssl']:
        if not os.path.exists(directory):
            os.makedirs(directory)
            app.logger.info(f"创建目录: {directory}")
    
    # 创建SSL上下文
    context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
    context.load_cert_chain('./ssl/cert.pem', './ssl/privkey.pem')
    
    try:
        app.logger.info("=== 启动HTTPS测试系统 ===")
        app.run(
            host='0.0.0.0', 
            port=app.config['HTTPS_PORT'], 
            ssl_context=context, 
            debug=False
        )
    except Exception as e:
        app.logger.critical(f"系统启动失败: {str(e)}")
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt, Inches
import os
app = Flask(__name__)

# Убедитесь, что у вас есть папка для хранения загруженных файлов
if not os.path.exists('downloads'):
    os.makedirs('downloads')
@app.route('/')
def index():
    return render_template('index.html')
@app.route('/gost')
def gost():
    return render_template('gost.html')
@app.route('/npk')
def npk():
    return render_template('npk.html')
@app.route('/rec')
def rec():
    return render_template('rec.html')
@app.route('/form', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        # Получение данных из формы
        topic = request.form['topic']
        fio = request.form['fio']
        class_name = request.form['class']
        char = request.form['char']
        UCH = request.form['UCH']
        post = request.form['post']
        year = request.form['year']
        intro = request.form['intro']
        relevance = request.form['relevance']
        purposes = request.form['purposes']
        tasks = request.form['tasks']
        object_study = request.form['object_study']
        subject = request.form['subject']
        hypothesis = request.form['hypothesis']
        research_methods = request.form['research_methods']
        pactical = request.form['pactical']
        chapter1 = request.form['chapter1']
        chapter2 = request.form['chapter2']
        conclusion = request.form['conclusion']
        litre = request.form['litre']

        # Создание документа на основе шаблона
        doc = Document('template.docx')
        style = doc.styles['Normal']
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        # Замена маркеров на данные из формы
        for para in doc.paragraphs:
            if '$TOPIC' in para.text:
                para.text = para.text.replace('$TOPIC', topic)
            if '$FIO' in para.text:
                para.text = para.text.replace('$FIO', fio)
            if '$CLASS' in para.text:
                para.text = para.text.replace('$CLASS', class_name)
            if '$CHAR' in para.text:
                para.text = para.text.replace('$CHAR', char)
            if '$UCH' in para.text:
                para.text = para.text.replace('$UCH', UCH)
            if '$POST' in para.text:
                para.text = para.text.replace('$POST', post)
            if '$YEAR' in para.text:
                para.text = para.text.replace('$YEAR', year)
            if '$INTRO' in para.text:
                para.text = para.text.replace('$INTRO', intro)
            if '$relevance' in para.text:
                para.text = para.text.replace('$relevance', relevance)
            if '$purposes' in para.text:
                para.text = para.text.replace('$purposes', purposes)
            if '$tasks' in para.text:
                pr=tasks.split('\n')
                # Добавляем нумерованный список
                for index, item in enumerate(pr, start=1):
                    para.text = para.text.replace('$tasks',f"{index}. {item}"+'\t$tasks')
                para.text = para.text.replace('$tasks','')
            if '$object_study' in para.text:
                para.text = para.text.replace('$object_study', object_study)
            if '$subject' in para.text:
                para.text = para.text.replace('$subject', subject)
            if '$hypothesis' in para.text:
                para.text = para.text.replace('$hypothesis', hypothesis)
            if '$research_methods' in para.text:
                para.text = para.text.replace('$research_methods', research_methods)
            if '$pactical' in para.text:
                para.text = para.text.replace('$pactical', pactical)
            if '$chapter1' in para.text:
                chapter1 = chapter1.replace("\n", "\t")
                para.text = para.text.replace('$chapter1', chapter1)     
            if '$chapter2' in para.text:
                chapter2 = chapter2.replace("\n", "\t")
                para.text = para.text.replace('$chapter2', chapter2)
            if '$conclusion' in para.text:
                conclusion = conclusion.replace("\n", "\t")
                para.text = para.text.replace('$conclusion', conclusion)
            if '$litre' in para.text:
                pr=litre.split('\n')
                # Добавляем нумерованный список
                for index, item in enumerate(pr, start=1):
                    para.text = para.text.replace('$litre',f"{index}. {item}"+'$litre')
                para.text = para.text.replace('$litre','')
        # Сохранение документа
        output_file = 'downloads/filled_document.docx'
        doc.save(output_file)

        return send_file(output_file, as_attachment=True)

    return render_template('form.html')

if __name__ == '__main__':
    app.run(debug=True)

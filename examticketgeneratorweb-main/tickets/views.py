import os
import logging
from django.shortcuts import render
from django.conf import settings
from django.http import FileResponse, HttpResponse
from .forms import TicketForm
from docx import Document
from .utils import read_questions_from_docx, simple_pairing, create_formatted_exam_docx, generate_latex, compile_pdf
from datetime import datetime
import openai
from .utils import generate_tickets_from_questions
from django.contrib.auth.decorators import login_required

logger = logging.getLogger(__name__)

openai.api_key = "openaikey"

def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except Exception as e:
        logger.exception("Ошибка чтения docx-файла")
        raise

def generate_questions_from_lecture(text, num_questions):
    prompt = (
        f"На основе следующего конспекта лекции, сгенерируй {num_questions} вопросов для экзамена.\n\n"
        f"{text}\n\n"
        "Верни вопросы списком, по одному в строке, без нумерации и лишних комментариев."
    )
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=1500
        )
        result_text = response['choices'][0]['message']['content']
        questions = [line.strip() for line in result_text.split('\n') if line.strip()]
        return questions
    except Exception as e:
        logger.exception("Ошибка при генерации вопросов через OpenAI")
        raise

@login_required()
def generate_tickets(request):
    error_message = None
    if request.method == 'POST':
        form = TicketForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                gen_source = form.cleaned_data['gen_source']
                uploaded_file = form.cleaned_data['file']


                media_dir = settings.MEDIA_ROOT
                os.makedirs(media_dir, exist_ok=True)
                tmp_path = os.path.join(media_dir, uploaded_file.name)
                try:
                    with open(tmp_path, 'wb+') as destination:
                        for chunk in uploaded_file.chunks():
                            destination.write(chunk)
                except Exception as e:
                    logger.exception("Ошибка при сохранении файла")
                    error_message = "Ошибка при загрузке файла. Проверьте, что файл не поврежден и не превышает размер."
                    raise


                direction = form.cleaned_data['direction']
                profile = form.cleaned_data['profile']
                department = form.cleaned_data['department']
                discipline = form.cleaned_data['discipline']

                num_tickets = form.cleaned_data.get('num_tickets', 10)
                tickets = []

                if gen_source == 'questions':
                    try:
                        tickets = generate_tickets_from_questions(tmp_path)
                        tickets = tickets[:num_tickets]
                    except Exception as e:
                        logger.exception("Ошибка при генерации билетов из вопросов")
                        error_message = "Не удалось обработать файл с вопросами. Убедитесь, что формат верный."
                        raise
                elif gen_source == 'lecture':
                    try:
                        lecture_text = extract_text_from_docx(tmp_path)
                        total_questions = num_tickets * 2
                        questions = generate_questions_from_lecture(lecture_text, total_questions)
                        tickets = [questions[i:i+2] for i in range(0, len(questions), 2)]
                    except Exception as e:
                        logger.exception("Ошибка при обработке конспекта лекции")
                        error_message = "Ошибка при обработке конспекта лекции или при генерации вопросов."
                        raise
                else:
                    error_message = "Неизвестный источник данных"
                    return HttpResponse(error_message, status=400)


                format_choice = form.cleaned_data.get('format', 'PDF')
                now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                output_dir = os.path.join(media_dir, f'tickets_{now}')
                os.makedirs(output_dir, exist_ok=True)

                if format_choice == 'DOCX':
                    try:
                        output_file = os.path.join(output_dir, 'tickets.docx')
                        create_formatted_exam_docx(tickets, output_file, discipline, "-", "-")
                        with open(output_file, 'rb') as f:
                            response = HttpResponse(
                                f.read(),
                                content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                            )
                            response['Content-Disposition'] = f'attachment; filename=tickets_{now}.docx'
                            return response
                    except Exception as e:
                        logger.exception("Ошибка при формировании DOCX")
                        error_message = "Не удалось сформировать документ DOCX."
                        raise
                else:  # PDF
                    try:
                        tex_file = os.path.join(output_dir, 'tickets.tex')
                        generate_latex(
                            tickets, tex_file,
                            direction=direction,
                            profile=profile,
                            department=department,
                            discipline=discipline,
                        )
                        compiled = compile_pdf(tex_file, output_dir)
                        if compiled:
                            pdf_file = os.path.join(output_dir, 'tickets.pdf')
                            with open(pdf_file, 'rb') as f:
                                response = HttpResponse(f.read(), content_type='application/pdf')
                                response['Content-Disposition'] = f'attachment; filename=tickets_{now}.pdf'
                                return response
                        else:
                            error_message = "Ошибка при компиляции PDF. Проверьте корректность данных."
                            return render(request, 'tickets/form.html', {'form': form, 'error_message': error_message})
                    except Exception as e:
                        logger.exception("Ошибка при формировании PDF")
                        error_message = "Не удалось сформировать PDF. Проверьте корректность данных."
                        return render(request, 'tickets/form.html', {'form': form, 'error_message': error_message})
            except Exception as e:
                if not error_message:
                    error_message = "Произошла непредвиденная ошибка. Попробуйте ещё раз или обратитесь к администратору."
                return render(request, 'tickets/form.html', {'form': form, 'error_message': error_message})
        else:
            error_message = "Проверьте правильность заполнения всех полей формы."
    else:
        form = TicketForm()

    return render(request, 'tickets/form.html', {'form': form, 'error_message': error_message})

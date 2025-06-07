from django import forms

class TicketForm(forms.Form):
    gen_source = forms.ChoiceField(
        choices=[('questions', 'Файл с готовыми вопросами'), ('lecture', 'Файл с лекцией')],
        widget=forms.RadioSelect,
        label='Источник генерации билетов'
    )
    file = forms.FileField(label="DOCX-файл (список вопросов или лекция)")
    direction = forms.CharField(label="Направление подготовки", max_length=255)
    profile = forms.CharField(label="Профили подготовки", max_length=255)
    department = forms.CharField(label="Кафедра", max_length=255)
    discipline = forms.CharField(label="Дисциплина", max_length=255)
    num_tickets = forms.IntegerField(label="Количество билетов", min_value=1, initial=10)
    format = forms.ChoiceField(
        choices=[('PDF', 'PDF'), ('DOCX', 'DOCX')],
        initial='PDF',
        widget=forms.RadioSelect,
        label='Формат билетов'
    )

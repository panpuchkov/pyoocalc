from django.contrib import admin

from .models import Question, Choice


class ChoiceInline(admin.TabularInline):
    model = Choice
    extra = 3


@admin.register(Choice)
class ChoiceAdmin(admin.ModelAdmin):
    list_filter = ('question__question_text', 'question__pub_date')
    list_display = ('question', 'choice_text', 'votes', )
# admin.site.register(Choice, ChoiceAdmin)


@admin.register(Question)
class QuestionAdmin(admin.ModelAdmin):
    # fields = ['pub_date', 'question_text']
    fieldsets = [
        (None,               {'fields': ['question_text']}),
        ('Date information', {'fields': ['pub_date']}),
    ]
    inlines = [ChoiceInline]
# admin.site.register(Question, QuestionAdmin)

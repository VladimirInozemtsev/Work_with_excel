from collections import defaultdict  # Импорт defaultdict для создания вложенных словарей с значениями по умолчанию.

from django.db import transaction  # Импорт transaction.atomic для управления транзакциями.
from rest_framework import serializers  # Импорт сериализаторов из Django REST Framework.

from apps.advances.models import Advance, Road, Taxi  # Импорт моделей Advance, Road и Taxi.
from apps.advances.validators import ExtraCosts  # Импорт валидатора для поля extra_costs.
from intra.tools.base_mixins import FlexMixin  # Импорт миксина для гибкости сериализации.
from intra.mixins import ChooseMixin, PartialTransferMixin  # Импорт миксинов для выбора и частичного переноса данных.
from intra.tools.different_functions import get_today  # Импорт функции для получения текущей даты.
from intra.tools.save_mixins import BulkSaveMixins, ListSaveSerializer  # Импорт миксина для массового сохранения и сериализатора для списков.


class BaseRoadSerializer(FlexMixin, ChooseMixin, PartialTransferMixin, serializers.ModelSerializer):
    # Сериализатор для модели Road с использованием миксинов.
    class Meta:
        model = Road  # Указание модели Road.
        fields = '__all__'  # Включение всех полей модели.


class BaseTaxiSerializer(FlexMixin, PartialTransferMixin, serializers.ModelSerializer):
    # Сериализатор для модели Taxi с использованием миксинов.
    class Meta:
        model = Taxi  # Указание модели Taxi.
        fields = '__all__'  # Включение всех полей модели.


class BaseAdvanceSerializer(FlexMixin, ChooseMixin, PartialTransferMixin, serializers.ModelSerializer):
    # Сериализатор для модели Advance с использованием миксинов.
    class Meta:
        model = Advance  # Указание модели Advance.
        fields = '__all__'  # Включение всех полей модели.


class AdvanceSerializer(BulkSaveMixins, BaseAdvanceSerializer):
    # Основной сериализатор для Advance с поддержкой массового сохранения.
    taxi = BaseTaxiSerializer(many=True, required=False)  # Вложенный сериализатор для списка такси.
    ticket = BaseRoadSerializer(many=True, required=False)  # Вложенный сериализатор для списка билетов.

    class Meta(BaseAdvanceSerializer.Meta):
        may_bound = False  # Дополнительное свойство, отключающее привязку.
        list_serializer_class = ListSaveSerializer  # Использование специального сериализатора для списков.
        fields = ['pk', 'assignment', 'status', 'start', 'stop', 'receive', 'balance', 'extra_costs',
                  'daily', 'check', 'taxi', 'ticket']  # Поля для сериализации.

    def to_representation(self, instance):
        # Преобразует объект модели в сериализованный вид.
        ret = super().to_representation(instance)  # Вызывает метод базового класса.
        return ret  # Возвращает результат.

    def validate_extra_costs(self, value):
        # Валидирует поле extra_costs.
        return ExtraCosts.model_validate(value).dict(exclude_none=True)  # Проверяет и форматирует данные.

    @transaction.atomic
    def update(self, instance, validated_data):
        # Обновляет объект в одной транзакции.
        return self.bulk_save(instance, validated_data)  # Использует метод массового сохранения.

    @transaction.atomic
    def create(self, instance, validated_data):
        # Создаёт объект в одной транзакции.
        return self.bulk_save(instance, validated_data)  # Использует метод массового сохранения.

    def data_for_create_excel(self, inst, instance, **kwargs):
        """Собирает данные для экспорта в Excel."""
        number = 1  # Счётчик расходов.
        praise_ticket = praise_taxi = praise_extra = 0  # Инициализация сумм расходов.
        data = defaultdict(lambda: defaultdict(dict))  # Вложенный словарь для данных.

        data['start'] = start = getattr(instance, "start", None)  # Дата начала командировки.
        data['stop'] = stop = getattr(instance, "stop", None)  # Дата окончания командировки.
        data['duration'] = (stop - start).days if start and stop else None  # Продолжительность.

        if daily := getattr(instance, "daily", 0):
            # Добавляет данные о суточных.
            data['expenses'][number]['prise'] = daily
            data['expenses'][number]['name'] = f"Суточные, дней: {data['duration']}"
            data['expenses'][number]['date'] = f"{start} - {stop}"

        if tickets := getattr(inst, "ticket", None):
            # Добавляет данные о билетах.
            for number, ticket in enumerate(tickets.all(), start=number+1):
                data['expenses'][number]['prise'] = ticket.praise
                praise_ticket += ticket.praise
                data['expenses'][number]['name'] = f'{ticket.initial_station} - {ticket.final_station}'
                data['expenses'][number]['date'] = ticket.date
                data['expenses'][number]['comment'] = ticket.comment
                data['expenses'][number]['type'] = ticket.type

        if taxi := getattr(inst, "taxi", None):
            # Добавляет данные о такси.
            for number, t in enumerate(taxi.all(), start=number+1):
                data['expenses'][number]['prise'] = t.praise
                praise_taxi += t.praise
                data['expenses'][number]['name'] = f'{t.initial_station} - {t.final_station}'
                data['expenses'][number]['date'] = t.date
                data['expenses'][number]['comment'] = t.comment

        if extra_costs := getattr(inst, "extra_costs", None):
            # Добавляет данные о дополнительных расходах.
            for number, costs in enumerate(extra_costs.values(), start=number+1):
                data['expenses'][number]['prise'] = costs.get('prise', 0)
                praise_extra += costs.get('prise', 0)
                data['expenses'][number]['name'] = costs.get('name', '')
                data['expenses'][number]['date'] = costs.get('date', '')
                data['expenses'][number]['comment'] = costs.get('comment', '')
                data['expenses'][number]['type'] = costs.get('type', '')

        data['receive'] = receive = getattr(instance, "receive", None)  # Сумма, полученная сотрудником.
        data['total_received'] = receive + praise_ticket  # Общая сумма полученного.
        data['total_spent'] = daily + praise_ticket + praise_taxi + praise_extra  # Общая сумма расходов.
        balance1 = getattr(instance, "balance", None)  # Баланс из модели.
        balance2 = data['total_spent'] - data['total_received']  # Пересчитанный баланс.

        if balance1 == balance2:
            data['balance'] = balance1  # Устанавливает баланс, если расчёты совпадают.
        else:
            raise ValueError(f'balance mismatch: {balance1} != {balance2}')  # Генерирует ошибку при расхождении балансов.

        place = instance.vizit.project.place.first()  # Данные о месте проекта.
        company = instance.vizit.project.to_contract.to_company  # Данные о компании.
        data['place'] = f'{place.country}, {place.sity}, {company.structure} {company.company}'  # Формирует строку места.
        data['assignment'] = getattr(instance, "assignment", None)  # Назначение командировки.
        data['status'] = getattr(instance, "status", None)  # Статус командировки.
        data['code'] = instance.vizit.project.code  # Код проекта.
        data['user'] = instance.vizit.user.get_full_name()  # Имя сотрудника.
        data['personal_number'] = instance.vizit.user.personal_number  # Табельный номер.
        data['position'] = instance.vizit.user.position.position  # Должность сотрудника.
        data['division'] = instance.vizit.user.position.division  # Подразделение сотрудника.
        data['date'] = get_today(format='DD.MM.YYYY', string=True)  # Текущая дата.

        return data  # Возвращает собранные данные.

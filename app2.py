# Обновленная функция для обработки данных
def process_data(cashin_data, cashout_data):
    currency_mapping = {
        "currency/e03f64a6-2225-11ed-0a80-073a00365127": "PLN",
        "currency/e15d9c47-2226-11ed-0a80-04b900364797": "USD",
        "currency/e1754d40-cc82-11ec-0a80-08ab00701a1e": "EUR"
    }
    payment_type_mapping = {
        "Card-in-showroom": "card",
        "Cash-in-showroom": "cash"
    }
    currency_totals = {cur: {"cash": 0, "card": 0, "count": 0, "test_count": 0} for cur in currency_mapping.values()}
    details = []

    for data, doc_type in [(cashin_data, "Приход"), (cashout_data, "Расход")]:
        for item in data.get("rows", []):
            currency_href = item["rate"]["currency"]["meta"]["href"]
            currency = next((v for k, v in currency_mapping.items() if k in currency_href), "UNKNOWN")
            amount = item["sum"] / 100  # Сумма хранится в копейках
            applicable = item["applicable"]
            
            # Определение типа платежа через атрибут PaymentType
            payment_type = "unknown"
            for attr in item.get("attributes", []):
                if attr["name"] == "PaymentType":
                    payment_type = payment_type_mapping.get(attr["value"]["name"], "unknown")
                elif attr["name"] == "test_order":
                    test_order = attr["value"]

            # Коррекция для расхода
            if doc_type == "Расход":
                amount = -amount

            # Обновляем остатки по валютам
            if applicable:
                if test_order:
                    currency_totals[currency]["test_count"] += 1
                else:
                    currency_totals[currency][payment_type] += amount
                    currency_totals[currency]["count"] += 1

            # Добавляем детали
            details.append({
                "date": item["moment"].split(" ")[0],
                "order_number": item["name"],
                "amount": amount,
                "currency": currency,
                "payment_type": payment_type,
                "doc_type": doc_type,
                "test_order": "yes" if test_order else "no",
                "comment": item.get("description", "")
            })
    
    return currency_totals, details

# Обновленная функция для создания Excel-отчёта
def create_excel_report(currency_totals, details):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Баланс по валютам"
    ws2 = wb.create_sheet("Детали ордеров")

    # Заполнение первой вкладки
    ws1.append(["Валюта", "Наличные", "Карта", "Количество документов", "Тестовые ордера"])
    for currency, data in currency_totals.items():
        ws1.append([currency, data["cash"], data["card"], data["count"], data["test_count"]])
        if data["cash"] < 0 or data["card"] < 0:
            for cell in ws1[-1]:
                cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

    # Заполнение второй вкладки
    ws2.append(["Дата", "Номер ордера", "Сумма", "Валюта", "Тип платежа", "Тип документа", "Test Order", "Комментарий"])
    for detail in details:
        ws2.append([
            detail["date"], detail["order_number"], detail["amount"],
            detail["currency"], detail["payment_type"], detail["doc_type"],
            detail["test_order"], detail["comment"]
        ])
    
    # Сохранение в память
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# В остальном код остается прежним

def generate_excel():
    cash_in_orders = fetch_orders("cashin")
    cash_out_orders = fetch_orders("cashout")
    
    cash_in_orders = filter_orders_by_date(cash_in_orders, start_date, end_date)
    cash_out_orders = filter_orders_by_date(cash_out_orders, start_date, end_date)
    
    income_totals, income_details = calculate_totals(cash_in_orders, is_income=True)
    expense_totals, expense_details = calculate_totals(cash_out_orders, is_income=False)
    
    all_details = sorted(income_details + expense_details, key=lambda x: x['date'])
    
    wb = Workbook()
    
    # Лист "Остатки по валютам"
    summary_sheet = wb.create_sheet(title="Остатки по валютам")
    summary_sheet.append(["Валюта", "Test_order cash", "Cash for Santander", "Карта", "Обработано документов"])
    
    summary_sheet.column_dimensions['B'].width = 15
    summary_sheet.column_dimensions['C'].width = 15
    summary_sheet.column_dimensions['E'].width = 20
    
    for currency, data in income_totals.items():
        net_cash_test = round(data['cash'] + expense_totals[currency]['cash'], 2)  # Placeholder, replace with logic for Test_order cash
        net_cash_santander = round(data['cash'] + expense_totals[currency]['cash'], 2)  # Placeholder, replace with logic for Cash for Santander
        net_card = round(data['card'] + expense_totals[currency]['card'], 2)
        net_count = data['count'] + expense_totals[currency]['count']
        summary_sheet.append([currency, net_cash_test, net_cash_santander, net_card, net_count])
    
    # Лист "Детали ордеров"
    details_sheet = wb.create_sheet(title="Детали ордеров")
    details_sheet.append(["Дата", "Номер ордера", "Test_order", "Cash, PLN", "PLN в кассе после данной операции", "Cash, USD", "Cash, EUR", "Card", "Currency", "Comment"])
    
    details_sheet.column_dimensions['A'].width = 15
    details_sheet.column_dimensions['C'].width = 12
    
    pln_total = 0
    
    for detail in all_details:
        cash_pln = round(detail['cash_pln'], 2) if detail['cash_pln'] else 0
        pln_total = round(pln_total + cash_pln, 2)

        # Определение значения Test_order
        test_order = any(attr['name'] == 'test_order' and attr['value'] for attr in detail.get('attributes', []))

        row = [
            detail['date'], 
            detail['name'], 
            test_order,
            cash_pln if cash_pln != "" else "", 
            pln_total,
            round(detail['cash_usd'], 2) if detail['cash_usd'] else "", 
            round(detail['cash_eur'], 2) if detail['cash_eur'] else "", 
            round(detail['card'], 2) if detail['card'] else "", 
            detail['currency'],
            detail['comment']
        ]
        details_sheet.append(row)
    
        current_row = details_sheet.max_row
        pln_total_cell = details_sheet.cell(row=current_row, column=5)
        
        if pln_total < 0:
            pln_total_cell.font = Font(color="FF0000")
        else:
            pln_total_cell.font = Font(color="000000")
    
    default_sheet = wb["Sheet"]
    wb.remove(default_sheet)
    
    # Сохранение файла в буфер
    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    
    return file_stream

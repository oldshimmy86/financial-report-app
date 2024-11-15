def calculate_totals(orders, is_income):
    totals = {'PLN': {'total': 0, 'cash': 0, 'card': 0, 'count': 0},
              'USD': {'total': 0, 'cash': 0, 'card': 0, 'count': 0},
              'EUR': {'total': 0, 'cash': 0, 'card': 0, 'count': 0}}
    order_details = []
    
    for order in orders:
        sum_value = round(order['sum'] / 100, 2)
        sum_value = sum_value if is_income else -sum_value
        currency_href = order['rate']['currency']['meta']['href']
        
        currency = None
        if "currency/e03f64a6-2225-11ed-0a80-073a00365127" in currency_href:
            currency = 'PLN'
        elif "currency/e15d9c47-2226-11ed-0a80-04b900364797" in currency_href:
            currency = 'USD'
        elif "currency/e1754d40-cc82-11ec-0a80-08ab00701a1e" in currency_href:
            currency = 'EUR'
        
        if currency:
            totals[currency]['total'] = round(totals[currency]['total'] + sum_value, 2)
            totals[currency]['count'] += 1

            payment_type = next((attr['value']['name'] for attr in order.get('attributes', []) if attr['name'] == "PaymentType"), "Unknown")
            if payment_type == "Cash-in-showroom":
                totals[currency]['cash'] = round(totals[currency]['cash'] + sum_value, 2)
            elif payment_type == "Card-in-showroom":
                totals[currency]['card'] = round(totals[currency]['card'] + sum_value, 2)

            comment = order.get("description", "")
            test_order = order.get("test_order", False)

            order_details.append({
                'date': order['moment'].split(' ')[0],
                'name': order['name'],
                'currency': currency,
                'cash_pln': sum_value if payment_type == "Cash-in-showroom" and currency == "PLN" else "",
                'cash_usd': sum_value if payment_type == "Cash-in-showroom" and currency == "USD" else "",
                'cash_eur': sum_value if payment_type == "Cash-in-showroom" and currency == "EUR" else "",
                'card': sum_value if payment_type == "Card-in-showroom" else "",
                'comment': comment,
                'test_order': test_order
            })
                
    return totals, order_details

def generate_excel():
    cash_in_orders = fetch_orders("cashin")
    cash_out_orders = fetch_orders("cashout")
    
    cash_in_orders = filter_orders_by_date(cash_in_orders, start_date, end_date)
    cash_out_orders = filter_orders_by_date(cash_out_orders, start_date, end_date)
    
    income_totals, income_details = calculate_totals(cash_in_orders, is_income=True)
    expense_totals, expense_details = calculate_totals(cash_out_orders, is_income=False)
    
    all_details = sorted(income_details + expense_details, key=lambda x: x['date'])
    
    wb = Workbook()
    
    summary_sheet = wb.create_sheet(title="Остатки по валютам")
    summary_sheet.append(["Валюта", "Наличные", "Карта", "Количество документов"])
    
    for currency, data in income_totals.items():
        net_cash = round(data['cash'] + expense_totals[currency]['cash'], 2)
        net_card = round(data['card'] + expense_totals[currency]['card'], 2)
        net_count = data['count'] + expense_totals[currency]['count']
        summary_sheet.append([currency, net_cash, net_card, net_count])
    
    details_sheet = wb.create_sheet(title="Детали ордеров")
    details_sheet.append(["Дата", "Номер ордера", "Cash, PLN", "PLN total", "Cash, USD", "Cash, EUR", "Card", "Currency", "Comment", "Test Order"])
    
    details_sheet.column_dimensions['A'].width = 15
    
    pln_total = 0
    
    for detail in all_details:
        cash_pln = round(detail['cash_pln'], 2) if detail['cash_pln'] else 0
        pln_total = round(pln_total + cash_pln, 2)

        row = [
            detail['date'], 
            detail['name'], 
            cash_pln if cash_pln != "" else "", 
            pln_total,
            round(detail['cash_usd'], 2) if detail['cash_usd'] else "", 
            round(detail['cash_eur'], 2) if detail['cash_eur'] else "", 
            round(detail['card'], 2) if detail['card'] else "", 
            detail['currency'],
            detail['comment'],
            detail['test_order']
        ]
        details_sheet.append(row)
    
        current_row = details_sheet.max_row
        pln_total_cell = details_sheet.cell(row=current_row, column=4)
        
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

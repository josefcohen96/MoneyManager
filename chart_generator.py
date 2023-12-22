from openpyxl.chart import PieChart, Reference, BarChart


def generate_pie_chart(sheet):
    labels = []
    amounts = []

    for row in sheet.iter_rows(min_row=2, max_col=4, max_row=sheet.max_row, values_only=True):
        label = row[1]
        amount = row[2]

        labels.append(label)
        amounts.append(amount)

    label_totals = {}

    for label, amount in zip(labels, amounts):
        if label in label_totals:
            label_totals[label] += amount
        else:
            label_totals[label] = amount

    sheet.cell(row=1, column=20, value="Total Labels")
    i = 0

    for label in label_totals:
        total_label = label
        total_amount = label_totals[label]
        sheet.cell(row=2 + i, column=20, value=total_label)
        sheet.cell(row=2 + i, column=21, value=total_amount)
        i += 1
        print("i:" + str(i))

    pie = PieChart()
    pie.title = "Pie Chart"
    # Create a Reference object for the data
    labels = Reference(sheet, min_col=20, min_row=2, max_row=2 + i - 1, max_col=20)
    data = Reference(sheet, min_col=21, min_row=1, max_row=2 + i - 1, max_col=21)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    # sheet.add_chart(pie, "J1")


def generate_bar_char(sheet):
    if "Bar Chart" in sheet:
        existing_chart = sheet["Bar Chart"]
        sheet.remove(existing_chart)

    labels = []
    amounts = []

    for row in sheet.iter_rows(min_row=2, max_col=4, max_row=sheet.max_row, values_only=True):
        label = row[1]
        amount = row[2]

        labels.append(label)
        amounts.append(amount)

    label_totals = {}

    for label, amount in zip(labels, amounts):
        if label in label_totals:
            label_totals[label] += amount
        else:
            label_totals[label] = amount

    sheet.cell(row=1, column=20, value="Total Labels")
    i = 0

    for label in label_totals:
        total_label = label
        total_amount = label_totals[label]
        sheet.cell(row=2 + i, column=20, value=total_label)
        sheet.cell(row=2 + i, column=21, value=total_amount)
        i += 1

    bar = BarChart()
    bar.title = "Pie Chart"
    # Create a Reference object for the data
    labels = Reference(sheet, min_col=20, min_row=2, max_row=2 + i - 1, max_col=20)
    data = Reference(sheet, min_col=21, min_row=1, max_row=2 + i - 1, max_col=21)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(labels)
    sheet.add_chart(bar, "J1")

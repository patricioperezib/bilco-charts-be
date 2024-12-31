from flask import Flask, request, jsonify, Response
import xlsxwriter
import io
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/process_summary_comparison', methods=['POST'])
def process_summary_comparison():
    try:
        # Receive the JSON payload
        data = request.get_json()
        entries = data.get('data', [])  # Yearly data
        consolidated_data = data.get('consolidated', [])  # Consolidated sheet data

        # Create the Excel workbook in memory
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # Formats
        bold_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
        title_format = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 12, 'bg_color': '#D9E1F2',
        })
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        left_align_format = workbook.add_format({'align': 'left', 'valign': 'vcenter'})
        decimal_format = workbook.add_format({'num_format': '0.00', 'align': 'center', 'valign': 'vcenter'})

        # Adjust column widths dynamically
        def adjust_column_widths(worksheet, data):
            col_widths = {}
            for row in data:
                for col_idx, cell in enumerate(row):
                    col_widths[col_idx] = max(col_widths.get(col_idx, 0), len(str(cell)) + 2)
            for col_idx, width in col_widths.items():
                worksheet.set_column(col_idx, col_idx, width)

        # Function to write combined scope data with left alignment for the first column
        def write_combined_data(worksheet, start_row, data, title_style=None):
            for row_num, row in enumerate(data):
                # Apply title styling for the first row in each section
                if row_num == 0 and title_style:
                    if len(row) > 1:  # Merge if more than one column
                        worksheet.merge_range(start_row, 0, start_row, len(row) - 1, row[0], title_style)
                    else:
                        worksheet.write(start_row, 0, row[0], title_style)
                else:
                    for col_num, cell in enumerate(row):
                        # Left-align the first column
                        if col_num == 0:
                            worksheet.write(start_row + row_num, col_num, cell, left_align_format)
                        else:
                            # Apply decimal format if the cell contains a number
                            if isinstance(cell, (int, float)):
                                worksheet.write(start_row + row_num, col_num, cell, decimal_format)
                            else:
                                worksheet.write(start_row + row_num, col_num, cell, center_format)

        def write_data(worksheet, start_row, data, title_style=None, data_style=center_format):
            for row_num, row in enumerate(data):
                if row_num == 0 and title_style:  # Apply title style for first row
                    # Write the header row with title_style (background color)
                    for col_num, cell in enumerate(row):
                        worksheet.write(start_row + row_num, col_num, cell, title_style)
                else:
                    for col_num, cell in enumerate(row):
                       # Apply decimal format if the cell contains a number
                        if isinstance(cell, (int, float)):
                            worksheet.write(start_row + row_num, col_num, cell, decimal_format)
                        else:
                            worksheet.write(start_row + row_num, col_num, cell, data_style)
        
        def write_data_2(worksheet, start_row, data, title_style=None, data_style=center_format):
            for row_num, row in enumerate(data):
                if row_num == 0 and title_style:  # Apply title style for first row
                    # Write the header row with title_style (background color)
                    for col_num, cell in enumerate(row):
                        worksheet.write(start_row + row_num, col_num, cell, title_style)
                else:
                    for col_num, cell in enumerate(row):
                       # Apply decimal format if the cell contains a number
                        if isinstance(cell, (int, float)):
                            worksheet.write(start_row + row_num, col_num, cell * 0.001, decimal_format)
                        else:
                            worksheet.write(start_row + row_num, col_num, cell, data_style)

        def write_consolidated_data(worksheet, start_row, data, title_style=None):
            for row_num, row in enumerate(data):
                if row_num == 0 and title_style:  # Apply title style for the header row
                    for col_num, cell in enumerate(row):
                        worksheet.write(start_row + row_num, col_num, cell, title_style)
                else:
                    for col_num, cell in enumerate(row):
                        # Align first column to the left for non-header rows
                        if col_num == 0:
                            worksheet.write(start_row + row_num, col_num, cell, left_align_format)
                        else:
                             # Apply decimal format if the cell contains a number
                            if isinstance(cell, (int, float)):
                                worksheet.write(start_row + row_num, col_num, cell, decimal_format)
                            else:
                                worksheet.write(start_row + row_num, col_num, cell, center_format)
                           
        # Add a function to generate the Scope 1 chart
        def add_scope1_chart(workbook, worksheet, consolidated_data):
            # Define rows and columns for the chart
            years = consolidated_data[0][1:]  # Row 1 (excluding the "Category" header)
            row_indices = []
            for idx, row in enumerate(consolidated_data):
                if row[0].startswith("1.1 ") or row[0].startswith("1.2 ") or row[0].startswith("1.3 ") or row[0].startswith("1.4 "):
                    row_indices.append(idx)

            # Ensure we have valid rows for the chart
            if not row_indices:
                return  # No valid categories found, skip the chart
            categories = [consolidated_data[row_idx][0] for row_idx in row_indices]  # Category names
            chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

            # Add data series to the chart
            for i, row_idx in enumerate(row_indices):
                chart.add_series({
                    'name': categories[i],
                    'categories': [worksheet.name, 0, 1, 0, len(years)],  # Years
                    'values': [worksheet.name, row_idx, 1, row_idx, len(years)],  # Data values
                })

            # Configure chart axes and title
            chart.set_title({'name': 'Scope 1'})
            chart.set_x_axis({'name': 'Years'})
            chart.set_y_axis({'name': 'Emissions (tCO2e)'})
            chart.set_style(11)

             # Configurar la leyenda con fuente más pequeña
            chart.set_legend({
                'font': {'size': 8},  # Ajusta el tamaño de fuente (puedes usar un valor menor o mayor)
                'position': 'right'   # Opcional: Posicionar la leyenda a la derecha del gráfico
            })


            # Insert column chart into the worksheet
            worksheet.insert_chart(len(consolidated_data) + 2, 0, chart, {'x_scale': 1.1, 'y_scale': 1.2})

            
        
         # Add a function to generate the Scope 1 chart
        def add_scope2_chart(workbook, worksheet, consolidated_data):
            # Define rows and columns for the chart
            years = consolidated_data[0][1:]  # Row 1 (excluding the "Category" header)
            # Find row indices for categories starting with "2.1" and "2.2"
            row_indices = []
            for idx, row in enumerate(consolidated_data):
                if row[0].startswith("2.1 ") or row[0].startswith("2.2 "):
                    row_indices.append(idx)
            
            # Ensure we have valid rows for the chart
            if not row_indices:
                return  # No 2.1 or 2.2 categories found, skip the chart
            categories = [consolidated_data[row_idx][0] for row_idx in row_indices]  # Category names
            chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

            # Add data series to the chart
            for i, row_idx in enumerate(row_indices):
                chart.add_series({
                    'name': categories[i],
                    'categories': [worksheet.name, 0, 1, 0, len(years)],  # Years
                    'values': [worksheet.name, row_idx, 1, row_idx, len(years)],  # Data values
                })

            # Configure chart axes and title
            chart.set_title({'name': 'Scope 2'})
            chart.set_x_axis({'name': 'Years'})
            chart.set_y_axis({'name': 'Emissions (tCO2e)'})
            chart.set_style(11)

            chart.set_legend({
                'font': {'size': 8},  # Ajusta el tamaño de fuente (puedes usar un valor menor o mayor)
                'position': 'right'   # Opcional: Posicionar la leyenda a la derecha del gráfico
            })

            # Insert chart into the worksheet
            worksheet.insert_chart(len(consolidated_data) + 20, 0, chart, {'x_scale': 1.1, 'y_scale': 1.2})

         # Add a function to generate the Scope 1 chart
        def add_scope3_chart(workbook, worksheet, consolidated_data):
            # Define rows and columns for the chart
            years = consolidated_data[0][1:]  # Row 1 (excluding the "Category" header)
           # Find row indices for categories starting with "2.1" and "2.2"
            row_indices = []
            for idx, row in enumerate(consolidated_data):
                if row[0].startswith("3.1 ") or row[0].startswith("3.2 ") or row[0].startswith("3.3 ") or row[0].startswith("3.4 ") or row[0].startswith("3.5 ") or row[0].startswith("3.6 ") or row[0].startswith("3.7 "):
                    row_indices.append(idx)
            
            # Ensure we have valid rows for the chart
            if not row_indices:
                return  # No 2.1 or 2.2 categories found, skip the chart
            categories = [consolidated_data[row_idx][0] for row_idx in row_indices]  # Category names
            categories = [consolidated_data[row_idx][0] for row_idx in row_indices]  # Category names
            chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

            # Add data series to the chart
            for i, row_idx in enumerate(row_indices):
                chart.add_series({
                    'name': categories[i],
                    'categories': [worksheet.name, 0, 1, 0, len(years)],  # Years
                    'values': [worksheet.name, row_idx, 1, row_idx, len(years)],  # Data values
                })

            # Configure chart axes and title
            chart.set_title({'name': 'Scope 3'})
            chart.set_x_axis({'name': 'Years'})
            chart.set_y_axis({'name': 'Emissions (tCO2e)'})
            chart.set_style(11)

            chart.set_legend({
                'font': {'size': 8},  # Ajusta el tamaño de fuente (puedes usar un valor menor o mayor)
                'position': 'right'   # Opcional: Posicionar la leyenda a la derecha del gráfico
            })

            # Insert chart into the worksheet
            worksheet.insert_chart(len(consolidated_data) + 2 + 36, 0, chart, {'x_scale': 1.1, 'y_scale': 1.2})
        
         # Add a function to generate the Scope 1 chart
        def add_scope1and2_chart(workbook, worksheet, consolidated_data):
            # Define rows and columns for the chart
            years = consolidated_data[0][1:]  # Row 1 (excluding the "Category" header)
           # Find row indices for categories starting with "2.1" and "2.2"
            row_indices = []
            for idx, row in enumerate(consolidated_data):
                if row[0].startswith("SCOPE 1") or row[0].startswith("SCOPE 2"):
                    row_indices.append(idx)
            
            # Ensure we have valid rows for the chart
            if not row_indices:
                return  # No 2.1 or 2.2 categories found, skip the chart
            categories = [consolidated_data[row_idx][0] for row_idx in row_indices]  # Category names
            categories = [consolidated_data[row_idx][0] for row_idx in row_indices]  # Category names
            chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

            # Add data series to the chart
            for i, row_idx in enumerate(row_indices):
                chart.add_series({
                    'name': categories[i],
                    'categories': [worksheet.name, 0, 1, 0, len(years)],  # Years
                    'values': [worksheet.name, row_idx, 1, row_idx, len(years)],  # Data values
                })

            # Configure chart axes and title
            chart.set_title({'name': 'Scope 1 & 2'})
            chart.set_x_axis({'name': 'Years'})
            chart.set_y_axis({'name': 'Emissions (tCO2e)'})
            chart.set_style(11)

            # Insert chart into the worksheet
            worksheet.insert_chart(0, 6, chart)

         # Add a function to generate the Scope 1 chart
        def add_scope1and2and3_chart(workbook, worksheet, consolidated_data):
            # Define rows and columns for the chart
            years = consolidated_data[0][1:]  # Row 1 (excluding the "Category" header)
           # Find row indices for categories starting with "2.1" and "2.2"
            row_indices = []
            for idx, row in enumerate(consolidated_data):
                if row[0].startswith("SCOPE 1") or row[0].startswith("SCOPE 2") or row[0].startswith("SCOPE 3"):
                    row_indices.append(idx)
            
            # Ensure we have valid rows for the chart
            if not row_indices:
                return  # No 2.1 or 2.2 categories found, skip the chart
            categories = [consolidated_data[row_idx][0] for row_idx in row_indices]  # Category names
            categories = [consolidated_data[row_idx][0] for row_idx in row_indices]  # Category names
            chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

            # Add data series to the chart
            for i, row_idx in enumerate(row_indices):
                chart.add_series({
                    'name': categories[i],
                    'categories': [worksheet.name, 0, 1, 0, len(years)],  # Years
                    'values': [worksheet.name, row_idx, 1, row_idx, len(years)],  # Data values
                })

            # Configure chart axes and title
            chart.set_title({'name': 'Scopes'})
            chart.set_x_axis({'name': 'Years'})
            chart.set_y_axis({'name': 'Emissions (tCO2e)'})
            chart.set_style(11)

            # Insert chart into the worksheet
            worksheet.insert_chart(16, 6, chart)

        # Add Consolidated Totals sheet
        consolidated_sheet = workbook.add_worksheet("Consolidated Totals")
        write_consolidated_data(consolidated_sheet, 0, consolidated_data, title_style=title_format)
        adjust_column_widths(consolidated_sheet, consolidated_data)

        # Add the Scope 1 chart
        add_scope1_chart(workbook, consolidated_sheet, consolidated_data)
        add_scope2_chart(workbook, consolidated_sheet, consolidated_data)
        add_scope3_chart(workbook, consolidated_sheet, consolidated_data)
        add_scope1and2_chart(workbook, consolidated_sheet, consolidated_data)
        add_scope1and2and3_chart(workbook, consolidated_sheet, consolidated_data)

          # Process each year's data
       # Crear gráficos de tipo pie directamente en las hojas de cada año
      # Crear gráficos de tipo pie en una nueva hoja para cada año
        for entry in entries:
            year = entry['year']
            summary = entry['summary']
            comparison = entry['comparison']
            combined = entry['combined']

            # Crear la hoja principal del año con los datos originales
            worksheet_data = workbook.add_worksheet(f"Year {year}")
            row_cursor = 0

            # Escribir los datos de resumen y combinados
            write_data(worksheet_data, row_cursor, summary, title_style=title_format)
            row_cursor += len(summary) + 1

              # Escribir los datos de resumen y combinados
            write_data(worksheet_data, row_cursor, comparison, title_style=title_format)
            row_cursor += len(summary) + 1

            write_combined_data(worksheet_data, row_cursor, combined, title_style=title_format)
            row_cursor += len(combined)

            

            # Ajustar anchos de columna
            adjust_column_widths(worksheet_data, summary + comparison + combined)

            # Definir las celdas específicas para Scope 1, 2 y 3 en la tabla de resumen
            scope_categories = ['Scope 1', 'Scope 2', 'Scope 3']
            category_start_row = 1  # Fila donde comienza Scope 1 (indexada desde 0)
            category_column = 0  # Columna de categorías
            value_column = 1  # Columna de valores (Emissionen (tCO2e))

            # Escribir las categorías y valores relevantes temporalmente si es necesario
            chart_categories_range = [worksheet_data.name, category_start_row + 1, category_column, category_start_row + 3, category_column]
            chart_values_range = [worksheet_data.name, category_start_row + 1, value_column, category_start_row + 3, value_column]
            print(chart_categories_range,chart_values_range)

            # Crear el gráfico pie
            pie_chart_scopes = workbook.add_chart({'type': 'pie'})
            pie_chart_scopes.add_series({
                'name': f'Total Emissions (Scopes 1, 2, 3) - {year}',
                'categories': chart_categories_range,
                'values': chart_values_range,
               'data_labels': {'value': True, 'percentage': True}
            })
            pie_chart_scopes.set_title({'name': f'Scope 1, 2, 3 ({year})'})

            table_start_row = 1
            table_start_col = 0
            table_width = 4  # Número de columnas ocupadas por la tabla
            chart_row = table_start_row  # Alinea el gráfico con el inicio de la tabla
            chart_col = table_start_col + table_width + 1  # Coloca el gráfico a la derecha de la tabla

            # Insertar el gráfico a la derecha de la tabla
            worksheet_data.insert_chart(chart_row, chart_col, pie_chart_scopes)

            # Crear una nueva hoja para los gráficos del año
            worksheet_chart = workbook.add_worksheet(f"Year Chart {year}")

            # Inicializar posición para escribir datos y gráficos en la hoja de gráficos
            temp_start_row = 0
            temp_start_col = 0

            # === LÓGICA PARA SCOPE 1 ===
            scope1_start = None
            scope1_end = None
            for row_idx, row in enumerate(combined):
                if "SCOPE 1 - Direkte Emissionen" in row[0] or "SCOPE 1 - Direct emissions" in row[0]:
                    scope1_start = row_idx + 1  # Primera fila después del encabezado
                if scope1_start and "GESAMT" in row[0]:  # Detectar el final del bloque Scope 1
                    scope1_end = row_idx
                    break
                if scope1_start and "TOTAL" in row[0]:  # Detectar el final del bloque Scope 1
                    scope1_end = row_idx
                    break

            if scope1_start and scope1_end:
                print("step 1")
                relevant_rows_scope1 = [
                    row_idx for row_idx in range(scope1_start, scope1_end)
                    if combined[row_idx][0].startswith(("1.1 ", "1.2 ", "1.3 ", "1.4 "))
                    and not combined[row_idx][0].startswith(("1.1.1", "1.1.2"))
                ]
                if relevant_rows_scope1:
                    print("step 2")
                    # Agregar encabezado "tCO₂" en la segunda columna
                    worksheet_chart.write(temp_start_row, temp_start_col + 1, "tCO₂", title_format)  # Fila inicial, columna 2
                    temp_start_row += 1  # Avanzar una fila para no sobreescribir el encabezado

                    for i, idx in enumerate(relevant_rows_scope1):
                        worksheet_chart.write(temp_start_row + i, temp_start_col, combined[idx][0])  # Categorías
                        worksheet_chart.write(temp_start_row + i, temp_start_col + 1, combined[idx][1],decimal_format)  # Valores
                    categories_range_scope1 = [worksheet_chart.name, temp_start_row, temp_start_col, temp_start_row + len(relevant_rows_scope1) - 1, temp_start_col]
                    values_range_scope1 = [worksheet_chart.name, temp_start_row, temp_start_col + 1, temp_start_row + len(relevant_rows_scope1) - 1, temp_start_col + 1]
                    pie_chart_scope1 = workbook.add_chart({'type': 'pie'})
                    pie_chart_scope1.add_series({
                        'name': f'Scope 1 Emissions for {year}',
                        'categories': categories_range_scope1,
                        'values': values_range_scope1,
                        'data_labels': {'value': True, 'percentage': True}
                    })
                    pie_chart_scope1.set_title({'name': f'Scope 1 ({year})'})
                    # Calcular la posición adecuada para insertar el gráfico
                    table_start_row = 1
                    table_start_col = 0
                    table_width = 4  # Número de columnas ocupadas por la tabla
                    chart_row = table_start_row  # Alinea el gráfico con el inicio de la tabla
                    chart_col = table_start_col + table_width + 1  # Coloca el gráfico a la derecha de la tabla

                    # Configurar la leyenda con fuente más pequeña
                    pie_chart_scope1.set_legend({
                        'font': {'size': 8},  # Ajusta el tamaño de fuente (puedes usar un valor menor o mayor)
                        'position': 'right'   # Opcional: Posicionar la leyenda a la derecha del gráfico
                    })

                    # Insertar el gráfico a la derecha de la tabla
                    worksheet_chart.insert_chart(chart_row, chart_col, pie_chart_scope1,  {'x_scale': 1.5, 'y_scale': 1})
                    # Ajustar anchos de columna para la tabla escrita
                    adjust_column_widths(worksheet_chart, [[combined[idx][0], combined[idx][1]] for idx in relevant_rows_scope1])

                    temp_start_row += len(relevant_rows_scope1) + 15  # Espacio para el siguiente bloque

            # === LÓGICA PARA SCOPE 2 ===
            scope2_start = None
            scope2_end = None
            for row_idx, row in enumerate(combined):
                if "SCOPE 2" in row[0]:
                    scope2_start = row_idx + 1  # Primera fila después del encabezado
                if scope2_start and "GESAMT" in row[0]:  # Detectar el final del bloque Scope 2
                    scope2_end = row_idx
                    break
                if scope2_start and "TOTAL" in row[0]:  # Detectar el final del bloque Scope 1
                    scope2_end = row_idx
                    break

            if scope2_start and scope2_end:
                relevant_rows_scope2 = [
                    row_idx for row_idx in range(scope2_start, scope2_end)
                    if combined[row_idx][0].startswith(("2.1", "2.2"))
                ]
                if relevant_rows_scope2:
                    # Agregar encabezado "tCO₂" en la segunda columna
                    worksheet_chart.write(temp_start_row, temp_start_col + 1, "tCO₂", title_format)  # Fila inicial, columna 2
                    temp_start_row += 1  # Avanzar una fila para no sobreescribir el encabezado
                    for i, idx in enumerate(relevant_rows_scope2):
                        worksheet_chart.write(temp_start_row + i, temp_start_col, combined[idx][0])  # Categorías
                        worksheet_chart.write(temp_start_row + i, temp_start_col + 1, combined[idx][1], decimal_format)  # Valores
                    categories_range_scope2 = [worksheet_chart.name, temp_start_row, temp_start_col, temp_start_row + len(relevant_rows_scope2) - 1, temp_start_col]
                    values_range_scope2 = [worksheet_chart.name, temp_start_row, temp_start_col + 1, temp_start_row + len(relevant_rows_scope2) - 1, temp_start_col + 1]
                    pie_chart_scope2 = workbook.add_chart({'type': 'pie'})
                    pie_chart_scope2.add_series({
                        'name': f'Scope 2 Emissions for {year}',
                        'categories': categories_range_scope2,
                        'values': values_range_scope2,
                        'data_labels': {'value': True, 'percentage': True}
                    })
                    pie_chart_scope2.set_title({'name': f'Scope 2 ({year})'})
                     # Calcular la posición adecuada para insertar el gráfico
                    table_start_row = 17
                    table_start_col = 0
                    table_width = 4  # Número de columnas ocupadas por la tabla
                    chart_row = table_start_row  # Alinea el gráfico con el inicio de la tabla
                    chart_col = table_start_col + table_width + 1  # Coloca el gráfico a la derecha de la tabla

                     # Configurar la leyenda con fuente más pequeña
                    pie_chart_scope2.set_legend({
                        'font': {'size': 8},  # Ajusta el tamaño de fuente (puedes usar un valor menor o mayor)
                        'position': 'right'   # Opcional: Posicionar la leyenda a la derecha del gráfico
                    })
                    # Insertar el gráfico a la derecha de la tabla
                    worksheet_chart.insert_chart(chart_row, chart_col, pie_chart_scope2, {'x_scale': 1.5, 'y_scale': 1})
                    temp_start_row += len(relevant_rows_scope2) + 15  # Espacio para el siguiente bloque

            # === LÓGICA PARA SCOPE 3 ===
            scope3_start = None
            scope3_end = None
            for row_idx, row in enumerate(combined):
                if "SCOPE 3" in row[0]:
                    scope3_start = row_idx + 1  # Primera fila después del encabezado
                if scope3_start and "GESAMT" in row[0]:  # Detectar el final del bloque Scope 3
                    scope3_end = row_idx
                    break
                if scope3_start and "TOTAL" in row[0]:  # Detectar el final del bloque Scope 1
                    scope3_end = row_idx
                    break

            if scope3_start and scope3_end:
                relevant_rows_scope3 = [
                    row_idx for row_idx in range(scope3_start, scope3_end)
                    if combined[row_idx][0].startswith(("3.1 ", "3.2 ", "3.3 ", "3.4 ", "3.5 ", "3.6 ", "3.7 "))
                ]
                if relevant_rows_scope3:
                     # Agregar encabezado "tCO₂" en la segunda columna
                    worksheet_chart.write(temp_start_row, temp_start_col + 1, "tCO₂", title_format)  # Fila inicial, columna 2
                    temp_start_row += 1  # Avanzar una fila para no sobreescribir el encabezado
                    for i, idx in enumerate(relevant_rows_scope3):
                        worksheet_chart.write(temp_start_row + i, temp_start_col, combined[idx][0])  # Categorías
                        worksheet_chart.write(temp_start_row + i, temp_start_col + 1, combined[idx][1], decimal_format)  # Valores
                    categories_range_scope3 = [worksheet_chart.name, temp_start_row, temp_start_col, temp_start_row + len(relevant_rows_scope3) - 1, temp_start_col]
                    values_range_scope3 = [worksheet_chart.name, temp_start_row, temp_start_col + 1, temp_start_row + len(relevant_rows_scope3) - 1, temp_start_col + 1]
                    pie_chart_scope3 = workbook.add_chart({'type': 'pie'})
                    pie_chart_scope3.add_series({
                        'name': f'Scope 3 Emissions for {year}',
                        'categories': categories_range_scope3,
                        'values': values_range_scope3,
                        'data_labels': {'value': True, 'percentage': True}
                    })
                    pie_chart_scope3.set_title({'name': f'Scope 3 ({year})'})
                      # Calcular la posición adecuada para insertar el gráfico
                    table_start_row = 33
                    table_start_col = 0
                    table_width = 4  # Número de columnas ocupadas por la tabla
                    chart_row = table_start_row  # Alinea el gráfico con el inicio de la tabla
                    chart_col = table_start_col + table_width + 1  # Coloca el gráfico a la derecha de la tabla

                     # Configurar la leyenda con fuente más pequeña
                    pie_chart_scope3.set_legend({
                        'font': {'size': 8},  # Ajusta el tamaño de fuente (puedes usar un valor menor o mayor)
                        'position': 'right'   # Opcional: Posicionar la leyenda a la derecha del gráfico
                    })
                    # Insertar el gráfico a la derecha de la tabla
                    worksheet_chart.insert_chart(chart_row, chart_col, pie_chart_scope3, {'x_scale': 1.5, 'y_scale': 1})
                    temp_start_row += len(relevant_rows_scope3) + 15  # Espacio para el siguiente bloque

      

        # Close the workbook
        workbook.close()
        output.seek(0)

        return Response(
            output.read(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment;filename=Summary_Comparison_Report.xlsx"}
        )
    except Exception as e:
        print("Error:", str(e))
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

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
                            worksheet.write(start_row + row_num, col_num, cell, center_format)

        def write_data(worksheet, start_row, data, title_style=None, data_style=center_format):
           for row_num, row in enumerate(data):
               if row_num == 0 and title_style:  # Apply title style for first row
                   # Write the header row with title_style (background color)
                   for col_num, cell in enumerate(row):
                       worksheet.write(start_row + row_num, col_num, cell, title_style)
               else:
                   for col_num, cell in enumerate(row):
                       worksheet.write(start_row + row_num, col_num, cell, data_style)

        def write_consolidated_data(worksheet, start_row, data, title_style=None):
           """
           Writes consolidated data with:
             - First row styled as title (centered, background color)
             - First column in non-header rows aligned to the left
             - Other cells centered
           """
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
                           worksheet.write(start_row + row_num, col_num, cell, center_format)
        # Process each year's data
        for entry in entries:
            year = entry['year']
            summary = entry['summary']
            comparison = entry['comparison']
            combined = entry['combined']

            worksheet = workbook.add_worksheet(f"Year {year}")
            row_cursor = 0

            # Write summary data
            write_data(worksheet, row_cursor, summary, title_style=title_format)
            row_cursor += len(summary) + 1

            # Write comparison data
            write_data(worksheet, row_cursor, comparison, title_style=title_format)
            row_cursor += len(comparison) + 1

            # Write combined scope data
            write_combined_data(worksheet, row_cursor, combined, title_style=title_format)
            row_cursor += len(combined)

            # Adjust column widths
            adjust_column_widths(worksheet, summary + comparison + combined)

        # Add Consolidated Totals sheet
        consolidated_sheet = workbook.add_worksheet("Consolidated Totals")
        write_consolidated_data(consolidated_sheet, 0, consolidated_data, title_style=title_format)
        adjust_column_widths(consolidated_sheet, consolidated_data)

        # Add charts for Scope 1, Scope 2, Scope 3
        def add_chart(sheet, scope_name, start_row, scope_rows, chart_position):
            chart = workbook.add_chart({'type': 'column'})
            for idx, row in enumerate(scope_rows):
                col_start = 1  # Assuming years start from column 1
                col_end = len(consolidated_data[0]) - 1
                chart.add_series({
                    'name':       f"={sheet.get_name()}!$A${row + 1}",
                    'categories': f"={sheet.get_name()}!$B$1:$H$1",
                    'values':     f"={sheet.get_name()}!$B${row + 1}:$H${row + 1}",
                    'fill': {'color': f"#{hex(0x3366CC + idx * 10000)[2:]}"},
                })
        chart.set_title({'name': scope_name})
        chart.set_x_axis({'name': 'Years'})
        chart.set_y_axis({'name': 't CO2e'})
        chart.set_legend({'position': 'bottom'})
        consolidated_sheet.insert_chart(chart_position, chart)

        # Rows for Scope 1, Scope 2, Scope 3
        scope_1_rows = [1, 2, 3, 4]
        scope_2_rows = [9, 10]
        scope_3_rows = [12, 13, 15, 17, 21, 24, 26]

        add_chart(consolidated_sheet, "Scope 1 - Direct Emissions", 1, scope_1_rows, 'J2')
        add_chart(consolidated_sheet, "Scope 2 - Indirect Emissions", 9, scope_2_rows, 'J17')
        add_chart(consolidated_sheet, "Scope 3 - Other Indirect Emissions", 12, scope_3_rows, 'J32')


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
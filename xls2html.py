# -*- coding: UTF-8 -*-
import xlrd


class Xls2html():
    def __init__(self, filepath,  sheet=None):
        self.filepath = filepath
        self.sheet = sheet

    def convert(self):
        self.wb = xlrd.open_workbook(
            filepath, formatting_info=True)
        self.ws = self.get_sheet()
        self.data = self.worksheet_to_data()
        self.html = self.render_data_to_html()

    def save(self, save_to):
        if self.html and save_to:
            output = open(save_to, 'w')
            output.write(self.html)

    def read(self):
        pass

    def get_sheet(self):
        if self.sheet is not None:
            if isinstance(self.sheet, int):
                ws = self.wb.sheet_by_index(self.sheet)
            else:
                ws = self.wb.sheet_by_name(self.sheet)
        else:
            ws = self.wb.sheet_by_index(0)
        return ws

    def render_data_to_html(self):
        html = '''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Title</title>
        <style>
            td {
                border: 1px solid black;
                height: 24px;
            };
            tr {
                height: 24px;
            }
        </style>
        </head>
        <body>
            %s
        </body>
        </html>
        '''
        return html % self.render_table()

    def render_attrs(self, attrs):
        if not attrs:
            return ''
        return ' '.join(["%s=%s" % a for a in sorted(attrs.items(), key=lambda a: a[0])])

    def render_inline_styles(self, styles):
        if not styles:
            return ''
        return ';'.join(
            ["%s: %s" % a for a in sorted(styles.items(), key=lambda a: a[0]) if a[1] is not None])

    def render_table(self):
        html = [
            '''
            <table style="border-collapse: collapse;empty-cells: show;" 
            border="0" cellspacing="0" cellpadding="0" width="100%">
            <colgroup>
            '''
        ]
        for col in self.data['cols']:
            html.append('<col {attrs} style="{styles}">'.format(
                attrs=self.render_attrs(col.get('attrs')),
                styles=self.render_inline_styles(col.get('style')),
            ))
        html.append('</colgroup>')

        for row in self.data['rows']:
            trow = ['<tr>']
            for cell in row:
                trow.append('<td {attrs_str} style="{styles_str}">{formatted_value}</td>'.format(
                    attrs_str=self.render_attrs(cell['attrs']),
                    styles_str=self.render_inline_styles(cell['style']),
                    **cell))

            trow.append('</tr>')
            html.append('\n'.join(trow))
        html.append('</table>')
        return '\n'.join(html)

    def coord(self, row, col):
        return "%s, %s" % (row, col)

    def get_merged_cell_map(self):
        merged_cell_map = {}
        excluded_cells = {}
        merged_cell_ranges = self.ws.merged_cells
        print(merged_cell_ranges)
        for rs in merged_cell_ranges:
            topRow = rs[0]
            bottomRow = rs[1] - 1
            topCol = rs[2]
            bottomCol = rs[3] - 1
            top_coord = self.coord(topRow, topCol)
            merged_cell_map[top_coord] = {
                'attrs': {
                    'colspan': bottomCol - topCol + 1,
                    'rowspan': bottomRow - topRow + 1,
                },
            }
            tempRow = topRow
            while tempRow <= bottomRow:
                tempCol = topCol
                while tempCol <= bottomCol:
                    coord = self.coord(tempRow, tempCol)
                    excluded_cells[coord] = 1
                    tempCol += 1
                tempRow += 1
            del excluded_cells[top_coord]
        return {"merged": merged_cell_map, "excluded": excluded_cells}

    def format_cell(self, cell):
        pass

    def get_styles_from_cell(self, cell, merged_cell_info):
        return {}

    def worksheet_to_data(self):
        data_list = []
        merged_cell = self.get_merged_cell_map()
        for row in range(self.ws.nrows):
            data_row = []
            data_list.append(data_row)
            for col in range(self.ws.ncols):
                coord = self.coord(row, col)
                cell = self.ws.cell(row, col)
                if coord in merged_cell["excluded"]:
                    continue
                cell_data = {
                    'row': row,
                    'col': col,
                    'value': self.ws.cell_value(row, col),
                    'formatted_value': self.ws.cell_value(row, col),
                    'attrs': {},
                    'style': {},
                }
                merged_cell_info = merged_cell["merged"].get(coord, {})
                if merged_cell_info:
                    cell_data['attrs'].update(merged_cell_info['attrs'])
                style = self.get_styles_from_cell(cell, merged_cell_info)
                cell_data['style'].update(style)
                data_row.append(cell_data)

        col_list = []
        for col in range(self.ws.ncols):
            width = self.ws.computed_column_width(col)
            col_width = width / 50
            col_list.append({
                'index': col,
                'style': {
                    "width": "{}px".format(col_width),
                }
            })

        return {'rows': data_list, 'cols': col_list}


if __name__ == "__main__":
    filepath = "现场复工流程-4.13.xls"
    x2h = Xls2html(filepath)
    x2h.convert()
    x2h.save("test.html")

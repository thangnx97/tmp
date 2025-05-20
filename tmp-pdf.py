from fpdf import FPDF
from time import time
class PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=10)
        self.t_wrap = 0
        self.t_add_page = 0
        self.t_add_row = 0
        self.t_get_line = 0

    def estimate_line_count(self, text, col_width):
        words = text.split()
        lines = 1
        line = ""

        for word in words:
            if pdf.get_string_width(line + " " + word) <= col_width:
                line += " " + word
            else:
                lines += 1
                line = word

        return lines
    
    def draw_row(self, idx, data, col_widths, wrap_indices):
        x_start = self.get_x()
        y_start = self.get_y()
        
        # First pass: measure height needed for each cell
        cell_heights = []
        t = time()
        for i, text in enumerate(data):
            w = col_widths[i]
            # if i in wrap_indices:
            tm = time()
            # lines = self.multi_cell(w, 5, text, border=0, split_only=True)
            lines = self.estimate_line_count(text, w)
            self.t_get_line+=time()-tm
            h = 5 * lines
            # else:
            #     h = 5
            cell_heights.append(h)
        
        row_height = max(cell_heights)
        t1 = time()
        self.t_wrap+=t1-t

        if y_start + row_height > self.page_break_trigger:
            self.add_page()
            x_start = self.l_margin
            y_start = self.get_y()

        t2 = time()
        self.t_add_page+=t2-t1
        # Second pass: draw each cell with proper height
        self.set_xy(x_start, y_start)
        for i, text in enumerate(data):
            w = col_widths[i]
            # if i in wrap_indices:
            x = self.get_x()
            y = self.get_y()
            if cell_heights[i] == row_height:
                self.multi_cell(w, 5, text, border=1)
            else:
                self.cell(w, row_height, text, border=1)
            # After multi_cell, cursor moves to next line. Move it back to top right of cell.
            self.set_xy(x + w, y)
            # else:
            #     self.multi_cell(w, row_height, text, border=1)  # no vertical centering
        self.ln(row_height)
        t3 = time()
        self.t_add_row+=t3-t2
pdf = PDF()
pdf.add_font('TimesNew', '', 'TIMES.TTF', uni=True)
pdf.add_font('TimesNew', 'B', 'TIMESBD.TTF', uni=True)
pdf.set_font('TimesNew', '', 9)
pdf.add_page()

col_widths = [15, 30, 60, 50, 40]
wrap_cols = {2, 3}  # which columns should wrap

headers = ["ID", "Code", "Description", "Details", "Price"]
pdf.draw_row(0, headers, col_widths, wrap_cols)

# Add some sample rows
rows = [
    ["1", "A01", "Mô tả dài cần xuống dòng nhiều lần để kiểm tra", "Chi tiết sản phẩm nằm trong phần này", "1000"],
    ["2", "B02", "Ngắn", "Thông tin chi tiết ngắn", "2000"],
    ["3", "C03", "Thông tin rất dài, mô tả nhiều dòng để thử khả năng xuống dòng của ô này", "Nội dung chi tiết hơn về sản phẩm", "3000"],
]

from time import time

st = time()

for i in range(100000):
    row = rows[i%3]
    row[0] = str(i)
    pdf.draw_row(i, row, col_widths, wrap_cols)
t = time()
print('Prepare time: ', t-st)
print('Wrap: ', pdf.t_wrap)
print('Lines: ', pdf.t_get_line)
print('Page: ', pdf.t_add_page)
print('Row: ', pdf.t_add_row)
pdf.output("final_wrapped_table.pdf")
print('Out: ', time() - t)

from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

COLOR_BLUE = RGBColor(0x00, 0x70, 0xC0)
COLOR_GREEN = RGBColor(0x00, 0xB0, 0x50)
COLOR_RED = RGBColor(0xC0, 0x00, 0x00)

FONT_NAME = "Sitka Display"
FONT_SIZE = Pt(8)

class WordService:

    def __init__(self, template_path): 
        self.document = Document(template_path) 

    def save(self, path):
        self.document.save(path)

    # --------------------------------------------------
    # INTERNAL HELPERS
    # --------------------------------------------------

    def _iter_paragraphs(self):
        for p in self.document.paragraphs:
            yield p
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p

    def _apply_style(self, run, color=None):
        run.font.name = FONT_NAME
        run.font.size = FONT_SIZE
        run.font.bold = False
        if color:
            run.font.color.rgb = color

    def _replace_placeholder(self, placeholder, text, color=None):
        found = False
        for paragraph in self._iter_paragraphs():
            full_text = "".join(run.text for run in paragraph.runs)
            if placeholder not in full_text:
                continue

            found = True
            paragraph.clear()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            before, after = full_text.split(placeholder, 1)

            if before:
                run_before = paragraph.add_run(before)
                self._apply_style(run_before)

            run_value = paragraph.add_run(text)
            self._apply_style(run_value, color)

            if after:
                run_after = paragraph.add_run(after)
                self._apply_style(run_after)
            
            break  # Stop after finding and replacing the first occurrence

    # --------------------------------------------------
    # DATE
    # --------------------------------------------------

    def replace_date(self, date_str):
        self._replace_placeholder("{{DATE}}", date_str)

    # --------------------------------------------------
    # TABLE 1
    # --------------------------------------------------

    def fill_table_1(self, data):
        # Check if dataframes are empty
        if data["most_active"].empty or data["best"].empty or data["worst"].empty:
            print("Warning: Some dataframes are empty, skipping table 1 fill")
            return
        
        for i in range(min(5, len(data["most_active"]), len(data["best"]), len(data["worst"]))):
            # Most Active
            if i < len(data["most_active"]):
                self._replace_placeholder(
                    f"{{{{MOST ACTIVE STOCK {i+1}}}}}",
                    data["most_active"].iloc[i]["name"]
                )

                self._replace_placeholder(
                    f"{{{{MAS MULTIPLE {i+1}}}}}",
                    f'{data["most_active"].iloc[i]["volume_multiple"]:.2f}x',
                    COLOR_BLUE
                )

            # Best Performer
            if i < len(data["best"]):
                self._replace_placeholder(
                    f"{{{{BEST PERFORMER {i+1}}}}}",
                    data["best"].iloc[i]["name"]
                )

                self._replace_placeholder(
                    f"{{{{INCREASE {i+1}}}}}",
                    f'+{data["best"].iloc[i]["variation"]*100:.2f}%',
                    COLOR_GREEN
                )

            # Worst Performer
            if i < len(data["worst"]):
                self._replace_placeholder(
                    f"{{{{WORST PERFORMER {i+1}}}}}",
                    data["worst"].iloc[i]["name"]
                )

                self._replace_placeholder(
                    f"{{{{DECREASE {i+1}}}}}",
                    f'{data["worst"].iloc[i]["variation"]*100:.2f}%',
                    COLOR_RED
                )

    # --------------------------------------------------
    # TABLE 2 & 3
    # --------------------------------------------------

    def fill_table_2(self, data):
        self._fill_market_tables(data, start_index=1)

    def fill_table_3(self, data):
        self._fill_market_tables(data, start_index=10)

    def _fill_market_tables(self, data, start_index):
        for i, row in enumerate(data):
            self._replace_placeholder(
                f"{{{{^{row['ticker']}}}}}",
                f"{row['close']:,.2f}".replace(",", " ")
            )

            idx = start_index + i
            if row["variation"] is None:
                self._replace_placeholder(f"{{{{MVT{idx}}}}}", "-")
            elif row["variation"] > 0:
                self._replace_placeholder(
                    f"{{{{MVT{idx}}}}}",
                    f'+{row["variation"]*100:.2f}%',
                    COLOR_GREEN
                )
            else:
                self._replace_placeholder(
                    f"{{{{MVT{idx}}}}}",
                    f'{row["variation"]*100:.2f}%',
                    COLOR_RED
                )
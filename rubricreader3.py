from docx.api import Document
import docx
from docx.enum.text import WD_COLOR_INDEX
from termcolor import colored

SHOW_EVERYTHING_INCLUDING_NON_HIGHLIGHTED = False

# for outputting pretty colors to the terminal
def convert_wd_color_index_to_termcolor(color_index):
    if (color_index == WD_COLOR_INDEX.BLUE):
        return "blue"
    if (color_index == WD_COLOR_INDEX.TURQUOISE):
        return "cyan"
    if (color_index == WD_COLOR_INDEX.GREEN):
        return "green"
    if (color_index == WD_COLOR_INDEX.BRIGHT_GREEN):
        return "light_green"
    if (color_index == WD_COLOR_INDEX.RED):
        return "red"
    if (color_index == WD_COLOR_INDEX.YELLOW):
        return "yellow"

    print("WARNING: Unrecognized color index: %s" % (color_index))
    return "white"

def calculateScoreFromHighlights(highlights):
    score = 0
    for h in highlights:
        score += h[1]
    return score

document = docx.Document('test2.docx')

print(colored("\n========= Found %d tables in the document ==========" % (len(document.tables)), "blue"))

# Print out summary of tables found in the document
for t, table in enumerate(document.tables):
    print(colored("Table %d has %d rows and %d columns" % (t, len(table.rows), len(table.columns)), "yellow"))    

# each of these lists will contain tuples of (text, score) which we'll later remove dupes using set
highlightedFoundationals = []
highlightedProficients = []
highlightedExemplarys = []

# Print out detailed contents of each table, along with what is highlighted
for t, table in enumerate(document.tables):
    print("\n\nTABLE %d:" % (t))
    for r, row in enumerate(table.rows):
        print("\n-------- Row %d --------" % (r))
        for c, cell in enumerate(row.cells):
            print()
            for p, paragraph in enumerate(cell.paragraphs):
                if (len(paragraph.runs) == 0):
                    continue
                # score will be the precentage of runs inside this paragraph that are highlighted
                numHighlightedRuns = 0
                text = paragraph.text
                for r2, run in enumerate(paragraph.runs):
                    if run.font.highlight_color is not None:
                        numHighlightedRuns += 1
                score = numHighlightedRuns / len(paragraph.runs)        # probably 1 or 0.5
                if (c == 1):
                    highlightedFoundationals.append((text, score))
                if (c == 2):
                    highlightedProficients.append((text, score))
                if (c == 3):
                    highlightedExemplarys.append((text, score))

                # just for nice output to the terminal
                for r2, run in enumerate(paragraph.runs):
                    if run.font.highlight_color is not None:
                        print(colored("*Table %d, Row %d, Cell %d, Paragraph %d, Run %d: %s" % (t, r, c, p, r2, run.text), convert_wd_color_index_to_termcolor(run.font.highlight_color)))
                    else:
                        if SHOW_EVERYTHING_INCLUDING_NON_HIGHLIGHTED:
                            print(" Table %d, Row %d, Cell %d, Paragraph %d, Run %d: %s" % (t, r, c, p, r2, run.text))


print("\n\n========= FOUNDATIONALS ==========")
highlightedFoundationals = list(set(highlightedFoundationals))
print(highlightedFoundationals)
print(calculateScoreFromHighlights(highlightedFoundationals))

print("\n\n========= PROFICIENTS ==========")
highlightedProficients = list(set(highlightedProficients))
# print(highlightedProficients)
print(calculateScoreFromHighlights(highlightedProficients))

print("\n\n========= EXEMPLARYS ==========")
highlightedExemplarys = list(set(highlightedExemplarys))
# print(highlightedExemplarys)
print(calculateScoreFromHighlights(highlightedExemplarys))

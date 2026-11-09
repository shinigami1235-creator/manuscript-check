# Manuscript Check

Checks a thesis against a school's thesis format and can make a corrected copy of a thesis Word file. It comes with the TULIBS template (Thammasat, CICM, English, Times New Roman, 2022) and can learn other formats from their templates. Everything runs on this computer with no internet once it is set up.

## Setup

Install Python 3.10 or newer, then run this once in this folder while connected to the internet:

    pip install -r requirements.txt

## Checking a thesis

Drag the thesis PDF onto `Check manuscript.bat`. Two files appear next to the thesis and the report opens:

- `<thesis> - format report.html` lists every problem with its page.
- `<thesis> - corrections.pdf` is ready to print. It opens with the full list, then shows every page that has a problem as a picture with each problem boxed and numbered, and the notes beside it. The numbers match the report.

Check the PDF the student saved from Word. A Word file also works on a computer with Microsoft Word since Word lays out the pages and refreshes the table of contents first. Without Word it falls back to LibreOffice, which can move page breaks and swap fonts, and the report says so.

## Fixing a thesis

Drag the thesis Word file (.docx) onto `Fix manuscript.bat`. The original file is not changed. These appear next to it:

- `<thesis> - fixed.docx` is the corrected copy.
- `<thesis> - fixed changes.txt` lists what was changed and what is left for the student.
- The report and corrections sheets for the corrected copy.

The fix covers margins, line spacing, text size and font, page numbers, headings, the table of contents, the lists of tables and figures, the order of the back pages and the spelling copied from the template. It does not touch the wording of the thesis, so a title that differs between pages and the advisor lines on the abstract page stay for the student.

Open the fixed copy in Word and answer Yes when Word asks to update fields, since the table of contents and the lists get their page numbers then. Save it as PDF and check the PDF.

## Adding another format

Drag the school's template PDF onto `Learn template.bat` and type a name for the format, such as Mahidol MSc thesis 2025. Put the Word copy of the template next to the PDF with the same name, since the margins are read from it. The program measures the template and saves a new rules file in the `rules` folder, then writes `<name> - format learned.txt` next to the template listing what it found: page size, font, text sizes, margins, line spacing, indent, page numbers, the order of the pages, chapter titles, cover lines and the fields on the abstract and biography pages.

Read that list against the school's format guide before using the format, since a template PDF cannot show everything. The bottom margin is assumed to be 2.54 cm unless the Word copy sets it, and typing errors in the template are copied as they are. Correct any number in the rules file with Notepad.

When the `rules` folder holds more than one format, Check and Fix ask which one to use.

The fix follows the format too, with one gap: page numbers are only fixed for formats that put them at the top right. For other positions the report still checks them.

## Fix and Check

Fix means the thesis differs from the template. Check means the template does not settle the point and the graduate office decides, such as the bottom margin, which the template file leaves at 0.49 cm.

## Changing the rules

Every measurement for a format sits in its file in the `rules` folder, such as `rules/tu-cicm-thesis-2022.json`: margins, font sizes, line spacing, page number style, section order, chapter titles and the template text that must be replaced. Edit a number there when the graduate office gives a different figure. Check and Fix both use the same file.

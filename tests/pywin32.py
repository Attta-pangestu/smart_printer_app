import win32com.client

class WordToPdfConverter:
    def __init__(self, word_file, output_pdf):
        self.word_file = word_file
        self.output_pdf = output_pdf
        self.word = win32com.client.Dispatch("Word.Application")
        self.word.Visible = False  # Make Word invisible

    def apply_page_setup(self, doc, orientation=2, paper_size=4, custom_margins=None, fit_to_one_page=True):
        """Apply page settings like orientation, paper size, and margins."""
        try:
            # Set page orientation (1 = Portrait, 2 = Landscape)
            doc.PageSetup.Orientation = orientation
            
            # Set paper size (1 = Letter, 4 = A4, 9 = Executive, etc.)
            doc.PageSetup.PaperSize = paper_size
            
            # Set custom margins (in inches)
            if custom_margins:
                doc.PageSetup.TopMargin = custom_margins.get("top", 1)
                doc.PageSetup.BottomMargin = custom_margins.get("bottom", 1)
                doc.PageSetup.LeftMargin = custom_margins.get("left", 1)
                doc.PageSetup.RightMargin = custom_margins.get("right", 1)
            
            # Fit to page
            if fit_to_one_page:
                # Disable zoom and apply fit-to-one-page settings
                doc.PageSetup.Zoom = False  # Disable zoom, let it scale automatically
                doc.PageSetup.FitToPagesWide = 1  # Fit to 1 page wide
                doc.PageSetup.FitToPagesTall = 1  # Fit to 1 page tall
        except Exception as e:
            print(f"Error during page setup: {e}")

    def convert_to_pdf(self):
        """Convert Word document to PDF with applied settings."""
        try:
            # Open the Word document
            doc = self.word.Documents.Open(self.word_file)

            # Apply page settings
            self.apply_page_setup(doc, orientation=2, paper_size=4, fit_to_one_page=True, custom_margins={"top": 1, "bottom": 1, "left": 1, "right": 1})

            # Export to PDF (FileFormat=17 corresponds to PDF format)
            doc.SaveAs(self.output_pdf, FileFormat=17)  # 17 is the PDF file format
            print(f"PDF saved at: {self.output_pdf}")
        
        except Exception as e:
            print(f"Error: {e}")
        finally:
            # Close the Word document and Word application
            doc.Close(False)
            self.word.Quit()

# Example usage
word_file = r"C:\Users\nbgmf\Downloads\Form Partisipasi Ekopesantren.docx"  # Path to your Word file
output_pdf = r"C:\Users\nbgmf\Downloads\Form Partisipasi Ekopesantren.pdf"  # Path to save the output PDF

# Create an instance of the converter
converter = WordToPdfConverter(word_file, output_pdf)

# Convert the Word file to PDF
converter.convert_to_pdf()

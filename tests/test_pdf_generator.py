import unittest
from src.pdf_generator import PDFGenerator

class TestPDFGenerator(unittest.TestCase):

    def setUp(self):
        self.pdf_generator = PDFGenerator()

    def test_merge_pdfs(self):
        # Test merging multiple PDFs into one
        pdf_files = ['test1.pdf', 'test2.pdf']
        output_file = 'merged_output.pdf'
        self.pdf_generator.merge_pdfs(pdf_files, output_file)
        # Check if the output file was created
        self.assertTrue(os.path.exists(output_file))

    def test_generate_pdf_filename(self):
        # Test the PDF filename generation
        company_name = "Test Company"
        date = "2023-10-01"
        expected_filename = "Test Company_2023-10-01.pdf"
        generated_filename = self.pdf_generator.generate_pdf_filename(company_name, date)
        self.assertEqual(generated_filename, expected_filename)

    def tearDown(self):
        # Clean up any created files after tests
        try:
            os.remove('merged_output.pdf')
        except OSError:
            pass

if __name__ == '__main__':
    unittest.main()
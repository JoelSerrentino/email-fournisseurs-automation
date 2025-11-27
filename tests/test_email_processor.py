import unittest
from unittest.mock import MagicMock, patch
from src.email_processor import EmailProcessor

class TestEmailProcessor(unittest.TestCase):

    @patch('src.outlook_handler.OutlookHandler')
    def setUp(self, MockOutlookHandler):
        self.mock_outlook_handler = MockOutlookHandler.return_value
        self.email_processor = EmailProcessor(self.mock_outlook_handler)

    def test_filter_emails(self):
        # Mock emails
        emails = [
            MagicMock(subject='Important: Supplier Invoice'),
            MagicMock(subject='Re: Your Order'),
            MagicMock(subject='Supplier Request for Quotation')
        ]
        keywords = ['Supplier', 'Invoice']
        filtered_emails = self.email_processor.filter_emails(emails, keywords)
        
        self.assertEqual(len(filtered_emails), 2)
        self.assertIn(emails[0], filtered_emails)
        self.assertIn(emails[2], filtered_emails)
        self.assertNotIn(emails[1], filtered_emails)

    def test_move_email(self):
        email = MagicMock()
        target_folder = 'Processed'
        self.email_processor.move_email(email, target_folder)
        
        self.mock_outlook_handler.move_email.assert_called_once_with(email, target_folder)

    def test_process_email(self):
        email = MagicMock(subject='Supplier Invoice from Company A')
        self.email_processor.move_email = MagicMock()
        self.email_processor.filter_emails = MagicMock(return_value=[email])
        
        self.email_processor.process_email(email)
        
        self.email_processor.move_email.assert_called_once_with(email, 'Processed')
        self.email_processor.filter_emails.assert_called_once()

if __name__ == '__main__':
    unittest.main()
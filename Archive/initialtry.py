#Imports for python base packages
import re
from datetime import datetime, timedelta
import logging

#Imports for packages that need to be installed
import win32com.client

class OutlookSecurityAgent:
    def __init__(self):
        """Initialize the Outlook Security Agent with connection to Outlook"""
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6)  # 6 represents the inbox

        # Configure logging
        logging.basicConfig(
            filename='outlook_security.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def check_phishing_indicators(self, email):
        """
        Check for common phishing indicators in an email
        Returns a list of detected indicators
        """
        indicators = []

        # Check sender domain mismatch
        sender = email.SenderEmailAddress.lower()
        display_name = email.SenderName.lower()
        if '@' in display_name and display_name != sender:
            indicators.append("Sender name/email mismatch")

        # Check for urgent language in subject
        urgent_words = ['urgent', 'immediate', 'action required', 'account suspended']
        if any(word in email.Subject.lower() for word in urgent_words):
            indicators.append("Urgent language in subject")

        # Check for suspicious links
        if email.HTMLBody:
            # Look for URLs that don't match their display text
            href_pattern = r'href=[\'"]?([^\'" >]+)'
            displayed_urls = re.findall(href_pattern, email.HTMLBody)
            for url in displayed_urls:
                if 'http' in url.lower():
                    if url.lower() not in email.HTMLBody.lower():
                        indicators.append("Mismatched URL display text")
                        break

        # Check for password or credential requests
        sensitive_words = ['password', 'login', 'credential', 'verify account']
        if any(word in email.Body.lower() for word in sensitive_words):
            indicators.append("Requests for sensitive information")

        return indicators

    def process_emails(self, days_back=1):
        """
        Process emails from the last specified number of days
        Move suspicious emails to a designated folder
        """
        try:
            # Create Security folder if it doesn't exist
            security_folder = None
            try:
                security_folder = self.namespace.GetDefaultFolder(6).Folders["Security Review"]
            except:
                security_folder = self.namespace.GetDefaultFolder(6).Folders.Add("Security Review")

            # Get recent emails
            restriction = "[ReceivedTime] >= '" + \
                (datetime.now() - timedelta(days=days_back)).strftime('%m/%d/%Y') + "'"
            emails = self.inbox.Items.Restrict(restriction)

            for email in emails:
                try:
                    indicators = self.check_phishing_indicators(email)

                    if indicators:
                        # Log the suspicious email
                        logging.info(f"Suspicious email detected:\nFrom: {email.SenderEmailAddress}\n" +
                                   f"Subject: {email.Subject}\nIndicators: {', '.join(indicators)}")

                        # Add warning to email subject
                        email.Subject = "[SUSPICIOUS] " + email.Subject

                        # Move to security review folder
                        email.Move(security_folder)

                except Exception as e:
                    logging.error(f"Error processing email: {str(e)}")

        except Exception as e:
            logging.error(f"Error in process_emails: {str(e)}")
            raise

    def get_security_stats(self):
        """Return statistics about processed emails"""
        security_folder = self.namespace.GetDefaultFolder(6).Folders["Security Review"]
        stats = {
            'total_flagged': len(security_folder.Items),
            'flagged_today': len(security_folder.Items.Restrict(
                "[ReceivedTime] >= '" + datetime.now().strftime('%m/%d/%Y') + "'"))
        }
        return stats

def main():
    """Main function to run the security agent"""
    agent = OutlookSecurityAgent()

    # Process last 24 hours of emails
    agent.process_emails(days_back=1)

    # Get and print statistics
    stats = agent.get_security_stats()
    print(f"Security Agent Report:")
    print(f"Total flagged emails: {stats['total_flagged']}")
    print(f"Flagged today: {stats['flagged_today']}")

if __name__ == "__main__":
    main()

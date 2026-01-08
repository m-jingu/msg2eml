#!/usr/bin/env python3
"""
MSG to EML Converter
A tool to convert Microsoft Office MSG files to EML format
"""

import argparse
import os
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate
from pathlib import Path

try:
    import extract_msg
except ImportError:
    print("Error: extract-msg library is not installed.", file=sys.stderr)
    print("Please install it with: pip install -r requirements.txt", file=sys.stderr)
    sys.exit(1)


def convert_msg_to_eml(msg_path, eml_path):
    """
    Convert MSG file to EML format
    
    Args:
        msg_path: Path to MSG file
        eml_path: Path to output EML file
    
    Returns:
        bool: True if conversion succeeds
    """
    try:
        # Read MSG file
        msg = extract_msg.Message(msg_path)
        
        # Create MIMEMultipart message
        eml_msg = MIMEMultipart('mixed')
        
        # Set basic headers
        if msg.sender:
            eml_msg['From'] = msg.sender
        if msg.to:
            eml_msg['To'] = msg.to
        if msg.cc:
            eml_msg['Cc'] = msg.cc
        if msg.bcc:
            eml_msg['Bcc'] = msg.bcc
        if msg.subject:
            eml_msg['Subject'] = msg.subject
        if msg.date:
            # Convert datetime object to RFC 2822 format string
            if isinstance(msg.date, datetime):
                eml_msg['Date'] = formatdate(timeval=msg.date.timestamp(), localtime=False)
            else:
                # Use as-is if already a string
                eml_msg['Date'] = str(msg.date)
        
        # Set body
        body_text = msg.body or ""
        body_html = msg.htmlBody or ""
        
        if body_html:
            # Use multipart/alternative if HTML body exists
            body_part = MIMEMultipart('alternative')
            if body_text:
                body_part.attach(MIMEText(body_text, 'plain', 'utf-8'))
            body_part.attach(MIMEText(body_html, 'html', 'utf-8'))
            eml_msg.attach(body_part)
        elif body_text:
            # Text body only
            eml_msg.attach(MIMEText(body_text, 'plain', 'utf-8'))
        
        # Process attachments
        if msg.attachments:
            for attachment in msg.attachments:
                try:
                    att_data = attachment.data
                    att_name = attachment.longFilename or attachment.shortFilename or "attachment"
                    
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(att_data)
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename="{att_name}"'
                    )
                    eml_msg.attach(part)
                except Exception as e:
                    print(f"Warning: Error processing attachment '{attachment.longFilename or attachment.shortFilename}': {e}", file=sys.stderr)
        
        # Save as EML file
        with open(eml_path, 'wb') as f:
            f.write(eml_msg.as_bytes())
        
        msg.close()
        return True
        
    except Exception as e:
        print(f"Error: Failed to convert MSG file '{msg_path}': {e}", file=sys.stderr)
        return False


def main():
    """Main function"""
    parser = argparse.ArgumentParser(
        description='Tool to convert MSG files to EML format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python msg2eml.py file1.msg
  python msg2eml.py file1.msg file2.msg file3.msg
        '''
    )
    parser.add_argument(
        'files',
        nargs='+',
        help='Path(s) to MSG file(s) to convert (multiple files allowed)'
    )
    
    args = parser.parse_args()
    
    success_count = 0
    error_count = 0
    
    for msg_file_path in args.files:
        # Convert file path to Path object
        msg_path = Path(msg_file_path)
        
        # Check if file exists
        if not msg_path.exists():
            print(f"Error: File not found: {msg_file_path}", file=sys.stderr)
            error_count += 1
            continue
        
        if not msg_path.is_file():
            print(f"Error: Not a file: {msg_file_path}", file=sys.stderr)
            error_count += 1
            continue
        
        # Generate output file path (same directory, change extension to .eml)
        eml_path = msg_path.with_suffix('.eml')
        
        # Execute conversion
        print(f"Converting: {msg_file_path} -> {eml_path}")
        if convert_msg_to_eml(str(msg_path), str(eml_path)):
            print(f"Completed: {eml_path}")
            success_count += 1
        else:
            error_count += 1
    
    # Display results
    print(f"\nConversion completed: {success_count} succeeded, {error_count} failed")
    
    # Return non-zero exit code if there were errors
    if error_count > 0:
        sys.exit(1)


if __name__ == '__main__':
    main()

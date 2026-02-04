"""
Utility functions for Excel file operations.
Reusable functions for S3 handling, path validation, and value checking.
"""

import os
import tempfile
from typing import Any, Tuple, Optional
from pathlib import Path
from urllib.parse import urlparse

try:
    import boto3
    from botocore.exceptions import ClientError, NoCredentialsError
    BOTO3_AVAILABLE = True
except ImportError:
    BOTO3_AVAILABLE = False


def is_s3_uri(path: str) -> bool:
    """
    Check if the path is an S3 URI.
    
    Args:
        path: File path to check
    
    Returns:
        True if path is an S3 URI, False otherwise
    """
    return path.startswith('s3://')


def download_from_s3(s3_uri: str, local_path: Optional[str] = None) -> str:
    """
    Download a file from S3 to a local temporary file.
    
    Args:
        s3_uri: S3 URI (e.g., 's3://bucket-name/path/to/file.xlsx')
        local_path: Optional local path to save the file. If None, creates a temp file.
    
    Returns:
        Path to the local file
    
    Raises:
        ImportError: If boto3 is not installed
        ValueError: If S3 URI is invalid
        RuntimeError: If AWS credentials are not configured or download fails
        FileNotFoundError: If file or bucket not found in S3
    """
    if not BOTO3_AVAILABLE:
        raise ImportError(
            "boto3 is required for S3 support. Install it with: pip install boto3"
        )
    
    parsed = urlparse(s3_uri)
    bucket_name = parsed.netloc
    s3_key = parsed.path.lstrip('/')
    
    if not bucket_name or not s3_key:
        raise ValueError(
            f"Invalid S3 URI: {s3_uri}. "
            f"Expected format: s3://bucket-name/path/to/file.xlsx"
        )
    
    try:
        s3_client = boto3.client('s3')
    except NoCredentialsError as e:
        raise RuntimeError(
            "AWS credentials not found. Please configure AWS credentials using:\n"
            "  - AWS CLI: aws configure\n"
            "  - Environment variables: AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY\n"
            "  - IAM role (if running on EC2)"
        ) from e
    
    if local_path is None:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        local_path = temp_file.name
        temp_file.close()
    
    try:
        s3_client.download_file(bucket_name, s3_key, local_path)
        return local_path
    except ClientError as e:
        error_code = e.response.get('Error', {}).get('Code', '')
        if error_code == 'NoSuchKey':
            raise FileNotFoundError(f"File not found in S3: {s3_uri}") from e
        if error_code == 'NoSuchBucket':
            raise FileNotFoundError(f"S3 bucket not found: {bucket_name}") from e
        raise RuntimeError(f"Error downloading from S3: {str(e)}") from e


def get_local_path(excel_path: str) -> Tuple[str, bool]:
    """
    Get local file path, downloading from S3 if necessary.
    
    Args:
        excel_path: Local file path or S3 URI
    
    Returns:
        Tuple of (local_file_path, is_temporary_file)
    
    Raises:
        FileNotFoundError: If local file doesn't exist or S3 file not found
    """
    if is_s3_uri(excel_path):
        local_path = download_from_s3(excel_path)
        return local_path, True
    
    if not Path(excel_path).exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    return excel_path, False


def is_blank_or_na(value: Any) -> bool:
    """
    Check if a value is blank, null, or N/A.
    
    Args:
        value: The value to check
    
    Returns:
        True if value is None, empty string, 'N/A', 'NA', 'null', or whitespace-only string
        False otherwise
    """
    if value is None:
        return True
    
    if isinstance(value, str):
        value_upper = value.strip().upper()
        blank_values = ['', 'N/A', 'NA', 'NULL', 'NONE', '#N/A', '#NA']
        if value_upper in blank_values:
            return True
    
    return False


def cleanup_temp_file(file_path: str) -> None:
    """
    Clean up a temporary file if it exists.
    
    Args:
        file_path: Path to the file to delete
    """
    if os.path.exists(file_path):
        try:
            os.unlink(file_path)
        except OSError:
            pass


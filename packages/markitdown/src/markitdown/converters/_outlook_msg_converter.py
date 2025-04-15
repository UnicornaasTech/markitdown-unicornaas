import sys
from typing import Any, Union, BinaryIO
from .._stream_info import StreamInfo
from .._base_converter import DocumentConverter, DocumentConverterResult
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE
from markdownify import markdownify as Markdownify


# Try loading optional (but in this case, required) dependencies
# Save reporting of any exceptions for later
_dependency_exc_info = None
olefile = None
try:
    import olefile  # type: ignore[no-redef]
except ImportError:
    # Preserve the error and stack trace for later
    _dependency_exc_info = sys.exc_info()

ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/vnd.ms-outlook",
]

ACCEPTED_FILE_EXTENSIONS = [".msg"]


class OutlookMsgConverter(DocumentConverter):
    """Converts Outlook .msg files to markdown by extracting email metadata and content.

    Uses the olefile package to parse the .msg file structure and extract:
    - Email headers (From, To, Subject)
    - Email body content
    """

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        # Check the extension and mimetype
        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        # Brute force, check if we have an OLE file
        cur_pos = file_stream.tell()
        try:
            if olefile and not olefile.isOleFile(file_stream):
                return False
        finally:
            file_stream.seek(cur_pos)

        # Brue force, check if it's an Outlook file
        try:
            if olefile is not None:
                msg = olefile.OleFileIO(file_stream)
                toc = "\n".join([str(stream) for stream in msg.listdir()])
                return (
                    "__properties_version1.0" in toc
                    and "__recip_version1.0_#00000000" in toc
                )
        except Exception as e:
            pass
        finally:
            file_stream.seek(cur_pos)

        return False

 

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        # Check: the dependencies
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".msg",
                    feature="outlook",
                )
            ) from _dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _dependency_exc_info[2]
            )

        assert (
            olefile is not None
        )  # If we made it this far, olefile should be available
        msg = olefile.OleFileIO(file_stream)

        # Extract email metadata
        md_content = "# Email Message\n\n"

        # Get headers
        headers = {
            # This got wrong fields from O365 MSG: "From": self._get_stream_data(msg, "__substg1.0_0C1F001F"),
            #Instead, get the name from '__substg1.0_0C1A001F, and then email address from __substg1.0_5D01001F
            "From": self._get_stream_data(msg, "__substg1.0_0C1A001F") + " <" + self._get_stream_data(msg, "__substg1.0_5D01001F") + ">",
            "To": self._get_stream_data(msg, "__substg1.0_0E04001F"),
            "Subject": self._get_stream_data(msg, "__substg1.0_0037001F"),
        }

        # Add headers to markdown
        for key, value in headers.items():
            if value:
                md_content += f"**{key}:** {value}\n"

        md_content += "\n## Content\n\n"


        # Get email body (plain text) - prefer it first
        body = self._get_stream_data(msg, "__substg1.0_1000001F")
        if body:
            md_content += body
        else:
            # Then, secondarily prefer HTML: try to get the stream, the one which has the HTML content
            if msg.exists("__substg1.0_10130102"):
                # Get raw binary data directly, as we need to decode it differently
                raw_data = msg.openstream("__substg1.0_10130102").read()
                html_content = self._process_html_stream(raw_data)
                if html_content and len(html_content) > 0:
                    md_content += html_content
        
        msg.close()

        return DocumentConverterResult(
            markdown=md_content.strip(),
            title=headers.get("Subject"),
        )

    def _get_stream_data(self, msg: Any, stream_path: str) -> Union[str, None]:
        """Helper to safely extract and decode stream data from the MSG file."""
        assert olefile is not None
        assert isinstance(
            msg, olefile.OleFileIO
        )  # Ensure msg is of the correct type (type hinting is not possible with the optional olefile package)

        try:
            if msg.exists(stream_path):
                data = msg.openstream(stream_path).read()

                # Try UTF-16 first (common for .msg files)
                try:
                    return self._strip_null_terminator(data.decode("utf-16-le"))
                except UnicodeDecodeError:
                    # Fall back to UTF-8
                    try:
                        return self._strip_null_terminator(data.decode("utf-8"))
                    except UnicodeDecodeError:
                        # Last resort - ignore errors
                        return self._strip_null_terminator(data.decode("utf-8", errors="ignore"))
        except Exception:
            pass
        return None
    

    def _strip_null_terminator(self, text: str) -> str:
        """
        Strips whitespace and removes trailing null character (\u0000) if present.
        (it seems that MSG files sometimes have a null in the end of the stream)
        
        Args:
            text: The string to strip
            
        Returns:
            The stripped string without trailing null character
        """
        if text and text.endswith('\u0000'):
            text = text[:-1]
        return text.strip()
    
    
    def _process_html_stream(self, raw_data: bytes) -> str:
        """Process raw HTML stream data by trying different encodings and validating the result.
          Unfortunately also the utf-16-le decoding works without errors even if the payload
          is actually iso-8859-1, so we have to check if the contents look like HTML and try again
          if the result does not look like HTML.

          We finally use markdownify to convert the HTML to markdown.

                  Args:
            raw_data: The raw binary data from the HTML stream
            
        Returns:
            The decoded HTML content as a markdown string
        """
        def is_valid_html(content: str) -> bool:
            # Look for common HTML markers that should appear in valid HTML
            markers = ['<html', '<body', '<head', '<div']
            content_lower = content.lower()
            return any(marker in content_lower for marker in markers)

        # Try UTF-16-LE first
        try:
            html = self._strip_null_terminator(raw_data.decode('utf-16-le'))
            if is_valid_html(html):
                return Markdownify(html)
            else:
                # If UTF-16-LE didn't give valid HTML, try ISO-8859-1
                html = self._strip_null_terminator(raw_data.decode('iso-8859-1'))
                if is_valid_html(html):
                    return Markdownify(html)
                else:
                    # Finally resort to utf-8, ignoring errors
                    html = self._strip_null_terminator(raw_data.decode('utf-8', errors='ignore'))
                    if is_valid_html(html):
                        return Markdownify(html)
        except UnicodeDecodeError:
            # If UTF fails, try ISO-8859-1
            try:
                html = self._strip_null_terminator(raw_data.decode('iso-8859-1'))
                if is_valid_html(html):
                    return Markdownify(html)
            except UnicodeDecodeError:
                pass
        return ""
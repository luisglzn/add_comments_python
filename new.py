import re
import numpy as np
from typing import Tuple, List, Optional

import re
import numpy as np
from typing import Tuple, Optional

def localize_substring_ignoring_separator(
    string: str, substring: str, separator: str = "<#>", case_sensitive: bool = True
) -> Optional[Tuple[int, int]]:
    """
    Locate the start and end indices of a substring within a string,
    ignoring occurrences of a specified separator.

    Args:
        string (str): The main string in which to search.
        substring (str): The substring to locate within the main string.
        separator (str): The separator to ignore in the search. Default is "<#>".
        case_sensitive (bool): Whether the search should be case-sensitive. Default is True.

    Returns:
        Optional[Tuple[int, int]]: A tuple containing the start and end indices of the located substring,
                                   or None if the substring is not found.
    """
    # Escape the separator for regex
    separator_pattern = re.escape(separator)

    # Normalize case if needed
    if not case_sensitive:
        string, substring = string.lower(), substring.lower()

    # Normalize the separator in both string and substring
    normalized_text = re.sub(separator_pattern, "", string)
    normalized_subtext = re.sub(separator_pattern, "", substring)

    # If the normalized text or substring is empty, return None
    if not normalized_text or not normalized_subtext:
        return None

    # Create an array to map indices from normalized_text back to the original string
    original_indices = np.arange(len(string))
    accumulated_diff = 0

    # Adjust indices to account for the differences due to separator normalization
    for match in re.finditer(separator_pattern, string):
        start_index = match.start() - accumulated_diff
        length_diff = len(match.group(0))
        original_indices[start_index:] += length_diff
        accumulated_diff += length_diff
    original_indices = original_indices.tolist()

    # Locate the substring in the normalized string
    normalized_location = normalized_text.find(normalized_subtext)
    if normalized_location == -1:
        return None

    # Map normalized indices back to the original indices
    original_location = (
        original_indices[normalized_location],
        original_indices[normalized_location + len(normalized_subtext) - 1] + 1,
    )
    return original_location


def localize_substring_all(
    string: str, substring: str, case_sensitive: bool = True
) -> List[Tuple[int, int]]:
    localizations = []
    shift = 0

    while True:
        loc = localize_substring(string, substring, case_sensitive)
        if loc is None:
            break
        si, ei = loc
        string = string[ei:]
        loc = si + shift, ei + shift
        shift += ei
        localizations.append(loc)

    return localizations


def localize_substring(
    string: str, substring: str, case_sensitive: bool = True
) -> Optional[Tuple[int, int]]:
    if not case_sensitive:
        string, substring = string.lower(), substring.lower()

    si = string.find(substring)
    if si == -1:
        return None

    ei = si + len(substring)
    return si, ei

def split_xml_by_elements(xml):
    """
    Split the XML content of a docx file by paragraphs.
    """

    paragraph_pattern = re.compile(
            r'<w:p(?:\s[^>]*)?>.*?</w:p>', re.DOTALL)
    return paragraph_pattern.findall(xml)

def build_txt(xml_paragraphs):
    def extract_text_from_paragraph(paragraph):
        # Buscar todas las etiquetas w:t y su contenido
        pattern = r'<w:t(?:\s+[^>]*)?>(.*?)</w:t>'
        matches = re.findall(pattern, paragraph)
        
        # Unir todos los textos encontrados
        return ''.join(matches)

    # Procesar cada párrafo y crear una nueva lista
    extracted_texts = []
    for paragraph in xml_paragraphs:
        text = extract_text_from_paragraph(paragraph)
        # Añadir el texto incluso si está vacío (para mantener saltos de línea)
        extracted_texts.append(text if text else ' ')
    
    return extracted_texts

def replace_tags(xml_list, separator):
    transformed_paragraphs = []
    tags = []
    
    for para in xml_list:
        tag_for_paragraph = []
        n_texts = para.count('<w:t>') + para.count('<w:t ')
        if n_texts == 0:
            tag_for_paragraph.append(para)
            para = separator
            transformed_paragraphs.append(para)
        elif n_texts == 1:
            pattern1 = re.compile(r'(<w:p(?:\s[^>]*)?>.*?<w:t(?:\s[^>]*)?>)(\s*)', re.DOTALL)
            match1 = pattern1.search(para)
            tag = match1.group(1) + match1.group(2)
            tag_for_paragraph.append(tag)
            para = pattern1.sub(separator, para)

            pattern2 = re.compile(r'(\s*)(</w:t[^>]*>.*?</w:p>)', re.DOTALL)
            match2 = pattern2.search(para)
            tag = match2.group(1) + match2.group(2)
            tag_for_paragraph.append(tag)
            para = pattern2.sub(separator, para)
            transformed_paragraphs.append(para)
        else:
            pattern1 = re.compile(r'(<w:p(?:\s[^>]*)?>.*?<w:t(?:\s[^>]*)?>)(\s*)', re.DOTALL)
            match1 = pattern1.search(para)
            tag = match1.group(1) + match1.group(2)
            tag_for_paragraph.append(tag)
            para = pattern1.sub(separator, para)

            for i in range(n_texts - 1):
                pattern2 = re.compile(r'(\s*)(</w:t[^>]*>.*?<w:t(?:\s[^>]*)?>)(\s*)', re.DOTALL)
                match2 = pattern2.search(para)
                tag = match2.group(1) + match2.group(2) + match2.group(3)
                tag_for_paragraph.append(tag)
                para = para.replace(match2.group(0), separator, 1)
            
            pattern3 = re.compile(r'(\s*)(</w:t[^>]*>.*?</w:p>)', re.DOTALL)
            match3 = pattern3.search(para)
            tag = match3.group(1) + match3.group(2)
            tag_for_paragraph.append(tag)
            para = pattern3.sub(separator, para)
            transformed_paragraphs.append(para)
        
        tags.append(tag_for_paragraph)
    return transformed_paragraphs, tags

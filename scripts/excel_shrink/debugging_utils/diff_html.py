
import xml.dom.minidom
import difflib
import webbrowser

def pretty_print_xml_bytes(bytes):
    string = bytes.decode('utf-8')
    ret = string.replace('<row', '\n<row')
    return ret

def diff_html_if_different(bytes1, bytes2):
    if bytes1 is None or bytes2 is None:
        return
    
    lines1 = pretty_print_xml_bytes(bytes1).splitlines()
    lines2 = pretty_print_xml_bytes(bytes2).splitlines()

    if lines1 == lines2:
        return

    html_diff = difflib.HtmlDiff().make_file(
        lines1, lines2,
        fromdesc='Original',
        todesc='Modified',
        context=True,
    )

    with open("diff_output.html", "w", encoding="utf-8") as f:
        f.write(html_diff)
        webbrowser.open("diff_output.html")


    return html_diff

    
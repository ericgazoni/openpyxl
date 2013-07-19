# Copyright (c) 2001-2011 Python Software Foundation
#
# License: PYTHON SOFTWARE FOUNDATION LICENSE VERSION 2
#          See http://www.opensource.org/licenses/Python-2.0 for full terms

import sys
from xml.sax.saxutils import XMLGenerator as _XMLGenerator


def quoteattr(data):
    data = data.replace("&", "&amp;")
    data = data.replace(">", "&gt;")
    data = data.replace("<", "&lt;")
    data = data.replace('\n', '&#10;')
    data = data.replace('\r', '&#13;')
    data = data.replace('\t', '&#9;')
    if '"' in data:
        if "'" in data:
            data = '"%s"' % data.replace('"', "&quot;")
        else:
            data = "'%s'" % data
    else:
        data = '"%s"' % data
    return data


if sys.version_info < (2, 5):

    try:
        from codecs import xmlcharrefreplace_errors
        _error_handling = "xmlcharrefreplace"
        del xmlcharrefreplace_errors
    except ImportError:
        _error_handling = "strict"

    class CompatXMLGenerator(_XMLGenerator):

        def _qname(self, name):
            """Builds a qualified name from a (ns_url, localname) pair"""
            if name[0]:
                # The name is in a non-empty namespace
                prefix = self._current_context[name[0]]
                if prefix:
                    # If it is not the default namespace, prepend the prefix
                    return prefix + ":" + name[1]
            # Return the unqualified name
            return name[1]

        def endElementNS(self, name, qname):
            self._write('</%s>' % self._qname(name))


else:
    CompatXMLGenerator = _XMLGenerator
    from xml.sax.saxutils import _error_handling


class XMLGenerator(CompatXMLGenerator):

    def startElementNS(self, name, qname, attrs):
        self._write('<' + self._qname(name))

        for (name, value) in attrs.items():
            self._write(' %s=%s' % (self._qname(name), quoteattr(value)))
        self._write('>')

    def _write(self, text):
        self._out.write(text.encode(self._encoding, _error_handling))


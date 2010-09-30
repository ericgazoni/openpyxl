# file openpyxl/shared/password_hasher.py

def hash_password(p_password = ''):
    """
    Create a password hash from a given string.

    This method is based on the algorithm provided by
    Daniel Rentz of OpenOffice and the PEAR package
    Spreadsheet_Excel_Writer by Xavier Noguer <xnoguer@rezebra.com>.
    """

    password = 0x0000
    i = 1

    for c in p_password:

        value = ord(c) << i
        rotated_bits = value >> 15
        value &= 0x7fff
        password ^= (value | rotated_bits)
        i += 1

    password ^= len(p_password)
    password ^= 0xCE4B

    return str(hex(password)).upper()[2:]


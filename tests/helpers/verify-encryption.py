"""Independent Office Agile interoperability check (msoffcrypto-tool 6.0.0).

Run through tests/encryption.test.ts with MSOFFCRYPTO_PYTHON pointing to a Python
installation containing msoffcrypto-tool. Password is supplied on stdin.
"""
import io
import sys
import zipfile

import msoffcrypto
import olefile
from msoffcrypto.exceptions import InvalidKeyError

path, plain_path = sys.argv[1:]
password = sys.stdin.read()
with open(path, 'rb') as source:
    document = msoffcrypto.OfficeFile(source)
    assert document.is_encrypted()
    try:
        document.load_key(password=password + '-wrong', verify_password=True)
    except InvalidKeyError:
        pass
    else:
        raise AssertionError('Wrong password was accepted')
    document.load_key(password=password, verify_password=True)
    output = io.BytesIO()
    document.decrypt(output, verify_integrity=True)
    with zipfile.ZipFile(output) as actual, zipfile.ZipFile(plain_path) as expected:
        assert actual.testzip() is None
        assert set(actual.namelist()) == set(expected.namelist())
        for name in actual.namelist():
            assert actual.read(name) == expected.read(name), name

# Modify ciphertext without changing the compound stream size.
with olefile.OleFileIO(path, write_mode=True) as container:
    payload = bytearray(container.openstream('EncryptedPackage').read())
    payload[16] ^= 1
    container.write_stream('EncryptedPackage', bytes(payload))
with open(path, 'rb') as source:
    document = msoffcrypto.OfficeFile(source)
    document.load_key(password=password, verify_password=True)
    try:
        document.decrypt(io.BytesIO(), verify_integrity=True)
    except InvalidKeyError:
        pass
    else:
        raise AssertionError('Modified ciphertext was accepted')
print('Independent decryption, password, integrity, and ZIP content checks passed')

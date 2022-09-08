/*
** Convert a COM binary buffer (bytes array / VT_ARRAY|VT_UI1) to a string,
** as UTF-8 or using the specified charset.
** 
** See HKEY_CLASSES_ROOT\MIME\Database\Charset for a list of supported charsets on your system.
** Warning: csISOLatin1, ISO_8859-1, ... are mapped to Windows 1252, not proper ISO/IEC 8859-1.
**
** Dependencies:
** - Microsoft ActiveX Data Objects ("ADODB.Stream")
*/ 
// For example, to explicitely convert using MS-DOS US codepage 437:
// var s = binaryToString(fileToBinary("axsh.ans"),"437");
// And if you want to handle SUB as an end-of-file like CP/M and MS-DOS does:
// s = s.replace(/\x1A[\s\S]*/,"")
//


function binaryToString (data/*, charset*/) {
	var adostm = new ActiveXObject("ADODB.Stream");
	var charset = arguments[1] || "utf-8";
	try {
		adostm.charset = charset;
	} catch(ex) {
		if (ex.number === -2146825287) { // Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.
			// charset requested isn't recognized
			var e = new TypeError("Unknown charset requested: "+ charset +".");
			e.description = e.message;
			throw e;
		} else {
			throw ex;
		}
	}
	adostm.type = 1 /*adTypeBinary*/;
	adostm.open();
	adostm.write(data);
	adostm.position = 0;
	adostm.type = 2 /*adTypeText*/;
	try {
		return adostm.readText();
	} catch(ex) {
		if (ex.number === -2147024809) { // The parameter is incorrect.
			var e = new TypeError("The data to be decoded contains an invalid character for charset "+ charset +".");
			e.description = e.message;
			throw e;
		} else {
			throw ex;
		}
	}
}

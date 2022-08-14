/*
** Convert a string to a COM stream object (IStream),
** as UTF-8 or using the specified charset.
** 
** See HKEY_CLASSES_ROOT\MIME\Database\Charset for a list of supported charsets on your system.
** 
** Dependencies:
** - Microsoft ActiveX Data Objects ("ADODB.Stream")
*/


function stringToBinary (text/*, charset*/) {
	var txtstm = new ActiveXObject("ADODB.Stream");
	txtstm.open();
	txtstm.type = 2 /*adTypeText*/;
	var charset = arguments[1] || "utf-8";
	try {
		txtstm.charset = charset;
	} catch(ex) {
		if (ex.number === -2146825287) { // Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.
			// charset requested isn't recognized
			var e = new TypeError("Unknown charset requested: "+ charset);
			e.description = e.message;
			throw e;
		} else {
			throw ex;
		}
	}
	var charset = txtstm.charset; // retrieve standardized charset name
	txtstm.writeText(String(text));
	
	// switch stream to binary mode and return its contents
	txtstm.position = 0;
	txtstm.type = 1 /*adTypeBinary*/;
	if (charset.toLowerCase() === "utf-8") {
		txtstm.position = 3; // skip UTF-8 BOM
	} else if (charset === "Unicode") {
		// This also handles "UTF-16" and "utf-16", they automatically get changed to "Unicode".
		txtstm.position = 2; // skip UTF-16 BOM
	} // No BOM is added for UTF-16BE
	return txtstm;
}

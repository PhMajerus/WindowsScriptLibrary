/*
** Convert a COM binary buffer (bytes array / VT_ARRAY|VT_UI1) contents to a
** COM stream object (IStream), using only components included with Windows.
** 
** Dependencies:
** - Microsoft ActiveX Data Objects ("ADODB.Stream")
*/


function binaryToStream (binary) {
	var adostm = new ActiveXObject("ADODB.Stream");
	adostm.type = 1 /*adTypeBinary*/;
	adostm.open();
	adostm.write(binary);
	adostm.position = 0;
	return adostm;
}

/*
** Automation DataView class for JScript
** 
** Provides a DataView-like interface to Automation binary buffers
** (VT_ARRAY|VT_UI1), which are similar to ArrayBuffers in COM/OLE
** Automation environments. It can also be used with ADODB.Stream objects by
** converting them to binary buffer using their .read() method.
** 
** Dependencies:
** - Microsoft ActiveX Data Objects ("ADODB.Stream")
** For getVarInt64, getVarUint64, setVarInt64, and setVarUint64 methods:
** - Majerus.net Automation Runtime library ("Majerus.Automation.*")
** 
** This class is designed to be familiar to developers used to ES6 DataView,
** or Majerus.net JSX DataView, but is not an exact drop-in replacement.
** While the ES6 DataView changes the buffer directly, this class copies the
** provided buffer and works on its internal copy. The getFullBuffer method
** can be used to get a new copy of its internal buffer after modifying it.
** If you need to access the data as an IStream, use the getFullStream method.
** In both cases, the returned value is a new copy, future changes through
** the view will not be applied until you request a new copy by calling the
** corresponding method again.
** 
** - Philippe Majerus, February 2020, updated May 2023 (Float16 support).
*/


// Note the object will work on a copy of the provided buffer, it will not modify the existing buffer.
// The .getFullBuffer() method can be used to retrive a copy of the modified buffer.
// To use an existing ADODB.Stream, use buffer.read() to pass a binary buffer copy.
function DataView(buffer /*, byteOffset, byteLength */) {
	if (this.constructor !== DataView)
		throw new Error("DataView is a constructor, it must be called with the \"new\" operator");
	
	// Store binary buffer in an ADO Stream.
	// It should not be tampered with for the other methods to work properly.
	this._stream = new ActiveXObject("ADODB.Stream");
	this._stream.type = 1 /*adTypeBinary*/;
	this._stream.open();
	this._stream.write(buffer);
	
	// The byteOffset property represents the offset (in bytes) of this view from the start.
	// In a standard JavaScript DataView it is a read-only accessor, here it's a simple
	// property, changing it will change the view window for following methods calls.
	this.byteOffset = arguments[1] || 0;
	if ((this.byteOffset < 0) || (this.byteOffset > this._stream.size)) {
		throw new Error("byteOffset is out of buffer bounds");
	}
	
	// The byteLength property represents the length (in bytes) of this view.
	// In a standard JavaScript DataView it is a read-only accessor, here it's a simple
	// property, changing it will change the view window for following methods calls.
	this.byteLength = arguments[2] || (this._stream.size-this.byteOffset);
	if ((this.byteOffset+this.byteLength) > this._stream.size) {
		throw new Error("byteLength is out of buffer bounds");
	}
	
	// Switch the stream to text so JavaScript can handle its
	// bytes as characters instead of COM byte arrays.
	this._stream.position = 0;
	this._stream.type = 2 /*adTypeText*/;
	this._stream.charset = "Windows-1252";
}


// Internal functions to write and read bytes to and from a DataView's internal stream object.
(function(){
	// Maps ADO Stream Windows-1252 characters to their byte values.
	var windows1252ReadMap = {
		'\0':0,'\x01':1,'\x02':2,'\x03':3,'\x04':4,'\x05':5,'\x06':6,'\x07':7,'\b':8,'\t':9,'\n':10,'\x0B':11,'\f':12,'\r':13,'\x0E':14,'\x0F':15,
		'\x10':16,'\x11':17,'\x12':18,'\x13':19,'\x14':20,'\x15':21,'\x16':22,'\x17':23,'\x18':24,'\x19':25,'\x1A':26,'\x1B':27,'\x1C':28,'\x1D':29,'\x1E':30,'\x1F':31,
		' ':32,'!':33,'\"':34,'#':35,'$':36,'%':37,'&':38,'\'':39,'(':40,')':41,'*':42,'+':43,',':44,'-':45,'.':46,'/':47,
		'0':48,'1':49,'2':50,'3':51,'4':52,'5':53,'6':54,'7':55,'8':56,'9':57,':':58,';':59,'<':60,'=':61,'>':62,'?':63,
		'@':64,'A':65,'B':66,'C':67,'D':68,'E':69,'F':70,'G':71,'H':72,'I':73,'J':74,'K':75,'L':76,'M':77,'N':78,'O':79,
		'P':80,'Q':81,'R':82,'S':83,'T':84,'U':85,'V':86,'W':87,'X':88,'Y':89,'Z':90,'[':91,'\\':92,']':93,'^':94,'_':95,
		'`':96,'a':97,'b':98,'c':99,'d':100,'e':101,'f':102,'g':103,'h':104,'i':105,'j':106,'k':107,'l':108,'m':109,'n':110,'o':111,
		'p':112,'q':113,'r':114,'s':115,'t':116,'u':117,'v':118,'w':119,'x':120,'y':121,'z':122,'{':123,'|':124,'}':125,'~':126,'\x7F':127,
		'\u20AC':128,'\x81':129,'\u201A':130,'\u0192':131,'\u201E':132,'\u2026':133,'\u2020':134,'\u2021':135,'\u02C6':136,'\u2030':137,'\u0160':138,'\u2039':139,'\u0152':140,'\x8D':141,'\u017D':142,'\x8F':143,
		'\x90':144,'\u2018':145,'\u2019':146,'\u201C':147,'\u201D':148,'\u2022':149,'\u2013':150,'\u2014':151,'\u02DC':152,'\u2122':153,'\u0161':154,'\u203A':155,'\u0153':156,'\x9D':157,'\u017E':158,'\u0178':159,
		'\xA0':160,'\xA1':161,'\xA2':162,'\xA3':163,'\xA4':164,'\xA5':165,'\xA6':166,'\xA7':167,'\xA8':168,'\xA9':169,'\xAA':170,'\xAB':171,'\xAC':172,'\xAD':173,'\xAE':174,'\xAF':175,
		'\xB0':176,'\xB1':177,'\xB2':178,'\xB3':179,'\xB4':180,'\xB5':181,'\xB6':182,'\xB7':183,'\xB8':184,'\xB9':185,'\xBA':186,'\xBB':187,'\xBC':188,'\xBD':189,'\xBE':190,'\xBF':191,
		'\xC0':192,'\xC1':193,'\xC2':194,'\xC3':195,'\xC4':196,'\xC5':197,'\xC6':198,'\xC7':199,'\xC8':200,'\xC9':201,'\xCA':202,'\xCB':203,'\xCC':204,'\xCD':205,'\xCE':206,'\xCF':207,
		'\xD0':208,'\xD1':209,'\xD2':210,'\xD3':211,'\xD4':212,'\xD5':213,'\xD6':214,'\xD7':215,'\xD8':216,'\xD9':217,'\xDA':218,'\xDB':219,'\xDC':220,'\xDD':221,'\xDE':222,'\xDF':223,
		'\xE0':224,'\xE1':225,'\xE2':226,'\xE3':227,'\xE4':228,'\xE5':229,'\xE6':230,'\xE7':231,'\xE8':232,'\xE9':233,'\xEA':234,'\xEB':235,'\xEC':236,'\xED':237,'\xEE':238,'\xEF':239,
		'\xF0':240,'\xF1':241,'\xF2':242,'\xF3':243,'\xF4':244,'\xF5':245,'\xF6':246,'\xF7':247,'\xF8':248,'\xF9':249,'\xFA':250,'\xFB':251,'\xFC':252,'\xFD':253,'\xFE':254,'\xFF':255
	};
	
	// Maps byte values to their ADO Stream Windows-1252 characters.
	var windows1252WriteMap = [
		'\0','\x01','\x02','\x03','\x04','\x05','\x06','\x07','\b','\t','\n','\x0B','\f','\r','\x0E','\x0F','\x10','\x11','\x12','\x13','\x14','\x15','\x16','\x17','\x18','\x19','\x1A','\x1B','\x1C','\x1D','\x1E','\x1F',
		' ','!','\"','#','$','%','&','\'','(',')','*','+',',','-','.','/','0','1','2','3','4','5','6','7','8','9',':',';','<','=','>','?',
		'@','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','[','\\',']','^','_',
		'`','a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','{','|','}','~','\x7F',
		'\x80','\x81','\x82','\x83','\x84','\x85','\x86','\x87','\x88','\x89','\x8A','\x8B','\x8C','\x8D','\x8E','\x8F','\x90','\x91','\x92','\x93','\x94','\x95','\x96','\x97','\x98','\x99','\x9A','\x9B','\x9C','\x9D','\x9E','\x9F',
		'\xA0','\xA1','\xA2','\xA3','\xA4','\xA5','\xA6','\xA7','\xA8','\xA9','\xAA','\xAB','\xAC','\xAD','\xAE','\xAF','\xB0','\xB1','\xB2','\xB3','\xB4','\xB5','\xB6','\xB7','\xB8','\xB9','\xBA','\xBB','\xBC','\xBD','\xBE','\xBF',
		'\xC0','\xC1','\xC2','\xC3','\xC4','\xC5','\xC6','\xC7','\xC8','\xC9','\xCA','\xCB','\xCC','\xCD','\xCE','\xCF','\xD0','\xD1','\xD2','\xD3','\xD4','\xD5','\xD6','\xD7','\xD8','\xD9','\xDA','\xDB','\xDC','\xDD','\xDE','\xDF',
		'\xE0','\xE1','\xE2','\xE3','\xE4','\xE5','\xE6','\xE7','\xE8','\xE9','\xEA','\xEB','\xEC','\xED','\xEE','\xEF','\xF0','\xF1','\xF2','\xF3','\xF4','\xF5','\xF6','\xF7','\xF8','\xF9','\xFA','\xFB','\xFC','\xFD','\xFE','\xFF'
	]; // We could use String.fromCharCode, as it tolerates surrogate half values, and Windows-1252 will decode C1 control characters, but this explicit map is cleaner as the opposite of the corresponding ReadMap.
	
	// Internal helper method to get an array of bytes from the view.
	DataView.prototype._getBytes = function (byteOffset, byteLength) {
		var pos = this.byteOffset + byteOffset;
		if ((pos < this.byteOffset) || (pos > this.byteOffset + this.byteLength)) {
			throw new Error("byteOffset is out of view bounds");
		}
		if ((pos + byteLength) > (this.byteOffset+this.byteLength)) {
			throw new Error("byteOffset plus size of data is out of view bounds");
		}
		this._stream.position = pos;
		var bytes = new Array();
		for (var i = 0; i<byteLength; i++) {
			var byte = windows1252ReadMap[this._stream.readText(1)];
			bytes.push(byte);
		}
		return bytes;
	};
	
	// Internal helper method to store an array of bytes into the view.
	DataView.prototype._setBytes = function (byteOffset, bytes) {
		var byteLength = bytes.length;
		var pos = this.byteOffset + byteOffset;
		if ((pos < this.byteOffset) || (pos > this.byteOffset + this.byteLength)) {
			throw new Error("byteOffset is out of view bounds");
		}
		if ((pos + byteLength) > (this.byteOffset+this.byteLength)) {
			throw new Error("byteOffset plus size of data is out of view bounds");
		}
		this._stream.position = pos;
		for (var i = 0; i < byteLength; i++) {
			var value = bytes[i];
			if (value < 0)
				value += 0x100;
			this._stream.writeText(windows1252WriteMap[value]);
		}
	};
})();


// Returns a copy of the view as a binary buffer (VT_ARRAY|VT_UI1).
DataView.prototype.getViewBuffer = function () {
	this._stream.position = 0;
	this._stream.type = 1 /*adTypeBinary*/;
	this._stream.position = this.byteOffset;
	var buffer = this._stream.read(this.byteLength);
	this._stream.position = 0;
	this._stream.type = 2 /*adTypeText*/;
	return buffer;
};

// Returns a copy of the view as an ADODB.Stream object that supports IStream.
DataView.prototype.getViewStream = function () {
	this._stream.position = 0;
	this._stream.type = 1 /*adTypeBinary*/;
	this._stream.position = this.byteOffset;
	
	var stream = new ActiveXObject("ADODB.Stream");
	stream.type = 1 /*adTypeBinary*/;
	stream.open();
	stream.write(this._stream.read(this.byteLength));
	stream.position = 0;
	
	this._stream.position = 0;
	this._stream.type = 2 /*adTypeText*/;
	
	return stream;
};

// Returns a copy of the full buffer as a binary buffer (VT_ARRAY|VT_UI1).
DataView.prototype.getFullBuffer = function () {
	this._stream.position = 0;
	this._stream.type = 1 /*adTypeBinary*/;
	var buffer = this._stream.read();
	this._stream.position = 0;
	this._stream.type = 2 /*adTypeText*/;
	return buffer;
};

// Returns a copy of the view as an ADODB.Stream object that supports IStream.
DataView.prototype.getFullStream = function () {
	var stream = new ActiveXObject("ADODB.Stream");
	stream.type = 1 /*adTypeBinary*/;
	stream.open();
	this._stream.copyTo(stream);
	stream.position = 0;
	return stream;
};


(function(){
	// Converts an array of 2 bytes into a half-precision floating-point value.
	function bytesToHalf(bytes) {
		// Sign (bit 15)
		var sign = (bytes[0] & 0x80) ? -1 : 1;
		// Exponent (bits 14-10)
		var exponent = ((bytes[0] & 0x7C) >> 2) - 15;
		// Fraction (bits 9-0)
		var mantissa = (bytes[1] / 256 + (bytes[0] & 0x03)) / 4;
		
		// Handle special values or combine result
		if (exponent === -15) {
			if (mantissa === 0) {
				return sign * 0;
			} else {
				// Subnormal number
				return sign * mantissa * Math.pow(2, -14);
			}
		} else if (exponent === 16) {
			if (mantissa === 0) {
				return sign * Infinity;
			} else {
				return NaN;
			}
		} else {
			// Normalized number, mantissa has implicit leading 1
			return sign * (1+mantissa) * Math.pow(2, exponent);
		}
	}
	
	// Gets a 16-bit float (half) at the specified byte offset from the start of the view.
	DataView.prototype.getFloat16 = function (byteOffset, littleEndian) {
		var bytes = this._getBytes(byteOffset, 2);
		if (littleEndian)
			bytes.reverse();
		return bytesToHalf(bytes);
	};
})();

(function(){
	// Converts an array of 4 bytes into a single-precision floating-point value.
	function bytesToSingle(bytes) {
		// Sign (bit 31)
		var sign = (bytes[0] & 0x80) ? -1 : 1;
		// Exponent (bits 30-23)
		var exponent = ((bytes[0] & 0x7F) << 1) + ((bytes[1] & 0x80) >> 7) - 127;
		// Fraction (bits 22-0)
		var mantissa = (((bytes[3]) / 256 + bytes[2]) / 256 + (bytes[1] & 0x7F)) / 128;
		
		// Handle special values or combine result
		if (exponent === -127) {
			if (mantissa === 0) {
				return sign * 0;
			} else {
				// Subnormal number
				return sign * mantissa * Math.pow(2, -126);
			}
		} else if (exponent === 128) {
			if (mantissa === 0) {
				return sign * Infinity;
			} else {
				return NaN;
			}
		} else {
			// Normalized number, mantissa has implicit leading 1
			return sign * (1+mantissa) * Math.pow(2, exponent);
		}
	}
	
	// Gets a 32-bit float (float/single) at the specified byte offset from the start of the view.
	DataView.prototype.getFloat32 = function (byteOffset, littleEndian) {
		var bytes = this._getBytes(byteOffset, 4);
		if (littleEndian)
			bytes.reverse();
		return bytesToSingle(bytes);
	};
})();

(function(){
	// Converts an array of 8 bytes into a double-precision floating-point value.
	function bytesToDouble(bytes) {
		// Sign (bit 63)
		var sign = (bytes[0] & 0x80) ? -1 : 1;
		// Exponent (bits 62-52)
		var exponent = ((bytes[0] & 0x7F) << 4) + ((bytes[1] & 0xF0) >> 4) - 1023;
		// Fraction (bits 51-0)
		var mantissa = ((((((bytes[7] / 256 + bytes[6]) / 256 + bytes[5]) / 256 + bytes[4]) / 256 + bytes[3]) / 256 + bytes[2]) / 256 + (bytes[1] & 0x0F)) / 16;
		
		// Handle special values or combine result
		if (exponent === -1023) {
			if (mantissa === 0) {
				return sign * 0;
			} else {
				// Subnormal number
				return sign * mantissa * Math.pow(2, -1022);
			}
		} else if (exponent === 1024) {
			if (mantissa === 0) {
				return sign * Infinity;
			} else {
				return NaN;
			}
		} else {
			// Normalized number, mantissa has implicit leading 1
			return sign * (1+mantissa) * Math.pow(2, exponent);
		}
	}
	
	// Gets a 64-bit float (double) at the specified byte offset from the start of the view.
	DataView.prototype.getFloat64 = function (byteOffset, littleEndian) {
		var bytes = this._getBytes(byteOffset, 8);
		if (littleEndian)
			bytes.reverse();
		return bytesToDouble(bytes);
	};
})();

// Gets a signed 8-bit integer (byte) at the specified byte offset from the start of the view.
DataView.prototype.getInt8 = function (byteOffset) {
	var bytes = this._getBytes(byteOffset, 1);
	var value = bytes[0];
	if (value >= 0x80)
		value -= 0x100;
	return value;
};

// Gets a signed 16-bit integer (short) at the specified byte offset from the start of the view.
DataView.prototype.getInt16 = function (byteOffset, littleEndian) {
	var bytes = this._getBytes(byteOffset, 2);
	if (littleEndian)
		bytes.reverse();
	var value = bytes.reduce(function(a,c){ return (a<<8)+c; });
	if (value >= 0x8000)
		value -= 0x10000;
	return value;
};

// Gets a signed 32-bit integer (long) at the specified byte offset from the start of the view.
DataView.prototype.getInt32 = function (byteOffset, littleEndian) {
	var bytes = this._getBytes(byteOffset, 4);
	if (littleEndian)
		bytes.reverse();
	var value = bytes.reduce(function(a,c){ return (a<<8)+c; });
	// JavaScript << operator automatically handles numbers as signed 32-bit, no fixup needed.
	return value;
};

// JScript does not include standard support for 64-bit integer values.
// Instead, we can create an Automation Int64 / LongLong variant (I8).
// They should be handled as opaque objects whose values you manipulate
// through the methods of a "Majerus.Automation.Int64" object.
DataView.prototype.getVarInt64 = function (byteOffset, littleEndian) {
	var VarInt64 = new ActiveXObject("Majerus.Automation.Int64");
	var bytes = this._getBytes(byteOffset, 8);
	if (littleEndian)
		bytes.reverse();
	var value = bytes.reduce(function(a,c){
		return VarInt64.Add(VarInt64.BitLShift(a,8), VarInt64.Convert(c));
	}, VarInt64.Zero);
	// Negative values are automatically properly handled by the BitLShift.
	return value;
};

// Gets an unsigned 8-bit integer (unsigned byte) at the specified byte offset from the start of the view.
DataView.prototype.getUint8 = function (byteOffset) {
	var bytes = this._getBytes(byteOffset, 1);
	return bytes[0];
};

// Gets an unsigned 16-bit integer (unsigned short) at the specified byte offset from the start of the view.
DataView.prototype.getUint16 = function (byteOffset, littleEndian) {
	var bytes = this._getBytes(byteOffset, 2);
	if (littleEndian)
		bytes.reverse();
	return bytes.reduce(function(a,c){ return (a<<8)+c; });
};

// Gets an unsigned 32-bit integer (unsigned long) at the specified byte offset from the start of the view.
DataView.prototype.getUint32 = function (byteOffset, littleEndian) {
	var bytes = this._getBytes(byteOffset, 4);
	if (littleEndian)
		bytes.reverse();
	var value = bytes.reduce(function(a,c){ return (a<<8)+c; });
	// For unsigned 32-bit, we must handle negative values explicitely,
	// JavaScript << operator automatically handles numbers as signed 32-bit, fix if negative.
	if (value < 0)
		value = 0x100000000 + value;
	return value;
};

// JScript does not include standard support for 64-bit integer values.
// Instead, we can create an Automation UInt64 / ULongLong variant (UI8).
// They should be handled as opaque objects whose values you manipulate
// through the methods of a "Majerus.Automation.UInt64" object.
DataView.prototype.getVarUint64 = function (byteOffset, littleEndian) {
	var VarUint64 = new ActiveXObject("Majerus.Automation.UInt64");
	var bytes = this._getBytes(byteOffset, 8);
	if (littleEndian)
		bytes.reverse();
	var value = bytes.reduce(function(a,c){
		return VarUint64.Add(VarUint64.BitLShift(a,8), VarUint64.Convert(c));
	}, VarUint64.Zero);
	return value;
};


(function(){
	// Converts a floating-point value into a half-precision array of 2 bytes.
	function halfToBytes (half) {
		if (isNaN(half)) {
			return [0x7F, 0xFF];
		} else if (half >= 65520) {
			return [0x7C, 0x00]; // Infinity
		} else if (half <= -65520) {
			return [0xFC, 0x00]; // -Infinity
		} else if (half === 0) {
			return [(1/half)<0 ? 0x80 : 0x00, 0x00];
		} else {
			var sign = half < 0 ? 1 : 0;
			if (sign)
				half = -half;
			// Compute the exponent (5 bits, signed).
			var exp = Math.floor(Math.log(half) / Math.LN2);
			var clampedexp = Math.max(-14, Math.min(exp, 15));
			var powexp = Math.pow(2, clampedexp);
			// Handle subnormals: leading digit is zero if exponent bits are all zero.
			var leading = exp <= -15 ? 0 : 1;
			if (!leading) {
				// Subnormal number
				clampedexp = -15;
			}
			// Compute 10 bits of mantissa, inverted to round toward zero.
			var mantissa = Math.round((leading - half / powexp) * 0x400);
			// reinvert mantissa and shift exponent for bias
			mantissa = -mantissa;
			clampedexp += 15; // bias
			return [(sign << 7) + (clampedexp << 2) + ((mantissa & 0x300) >> 8), mantissa & 0x0FF];
		}
	}
	
	// Stores a 16-bit float (half) value at the specified byte offset from the start of the view.
	DataView.prototype.setFloat16 = function (byteOffset, value, littleEndian) {
		var bytes = halfToBytes(value);
		if (littleEndian)
			bytes.reverse();
		this._setBytes(byteOffset, bytes);
	};
})();

(function(){
	// Converts a floating-point value into a single-precision array of 4 bytes.
	function singleToBytes (single) {
		if (isNaN(single)) {
			return [0x7F, 0xFF, 0xFF, 0xFF];
		} else if (single === Infinity) {
			return [0x7F, 0x80, 0x00, 0x00];
		} else if (single === -Infinity) {
			return [0xFF, 0x80, 0x00, 0x00];
		} else if (single === 0) {
			return [(1/single)<0 ? 0x80 : 0x00, 0x00, 0x00, 0x00];
		} else {
			var sign = single < 0 ? 1 : 0;
			if (sign)
				single = -single;
			// Compute the exponent (8 bits, signed).
			var exp = Math.floor(Math.log(single) / Math.LN2);
			var clampedexp = Math.max(-126, Math.min(exp, 127));
			var powexp = Math.pow(2, clampedexp);
			// Handle subnormals: leading digit is zero if exponent bits are all zero.
			var leading = exp <= -127 ? 0 : 1;
			if (!leading) {
				// Subnormal number
				clampedexp = -127;
			}
			// Compute 23 bits of mantissa, inverted to round toward zero.
			var mantissa = Math.round((leading - single / powexp) * 0x800000);
			if (exp > 0 && mantissa <= -0x800000) {
				// Number is too large for float32, follow Math.fround and ES6
				// DataView#setFloat32 behavior, returning bytes for +/-Infinity.
				return [(sign << 7) + 0x7F, 0x80, 0x00, 0x00];
			}
			// reinvert mantissa and shift exponent for bias
			mantissa = -mantissa;
			clampedexp += 127; // bias
			return [(sign << 7) + ((clampedexp & 0xFE) >>> 1), ((clampedexp & 0x01) << 7) + ((mantissa & 0x7F0000) >>> 16), (mantissa & 0x00FF00) >>> 8, mantissa & 0x0000FF];
		}
	}
	
	// Stores a 32-bit float (float/single) value at the specified byte offset from the start of the view.
	DataView.prototype.setFloat32 = function (byteOffset, value, littleEndian) {
		var bytes = singleToBytes(value);
		if (littleEndian)
			bytes.reverse();
		this._setBytes(byteOffset, bytes);
	};
})();

(function(){
	// Converts a floating-point value into a double-precision array of 8 bytes.
	function doubleToBytes (double) {
		if (isNaN(double)) {
			return [0x7F, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF];
		} else if (double === Infinity) {
			return [0x7F, 0xF0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
		} else if (double === -Infinity) {
			return [0xFF, 0xF0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
		} else if (double === 0) {
			return [(1/double)<0 ? 0x80 : 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
		} else {
			var mantissa = double;
			var sign = mantissa < 0 ? 1 : 0;
			if (sign)
				mantissa = -mantissa;
			var exp = 0;
			// Handle subnormals
			if (mantissa < Math.pow(2,-1022)) {
				// Subnormal number
				mantissa /= Math.pow(2,-1023-52);
			} else {
				// Compute the exponent (11 bits, signed).
				
				// Conceptually, we're performing the same as the code commented-out below...
				//exp = Math.floor(Math.log(mantissa) / Math.LN2);
				// clamp exponent
				//exp = Math.max(-1022, Math.min(exp, 1023));
				// compute mantissa value
				//mantissa /= Math.pow(2,exp-52);
				//exp += 1023; // bias
				
				// ... But we have to perform it manually to avoid rounding
				// errors with some border values such as Number.MAX_SAFE_INTEGER.
				while ((mantissa > Number.MAX_SAFE_INTEGER) && (exp < 1023)) {
					mantissa /= 2;
					exp++;
				}
				while ((mantissa <= (Number.MAX_SAFE_INTEGER/2)) && (exp > -1022)) {
					mantissa *= 2;
					exp--;
				}
				exp += 52 + 1023; // offset for mantissa size + bias
			}
			// Split mantissa into two 32-bit words, so bitwise operators can be used.
			var hi = (mantissa / 4294967296) | 0;
			var lo = mantissa & 0xFFFFFFFF;
			return [(sign << 7) + ((exp & 0x7F0) >>> 4), ((exp & 0x0F) << 4) + ((hi >>> 16) & 0x0F),
				(hi >>> 8) & 0xFF, hi & 0xFF, (lo >>> 24) & 0xFF, (lo >>> 16) & 0xFF, (lo >>> 8) & 0xFF, lo & 0xFF];
		}
	}
	
	// Stores a 64-bit float (double) value at the specified byte offset from the start of the view.
	DataView.prototype.setFloat64 = function (byteOffset, value, littleEndian) {
		var bytes = doubleToBytes(value);
		if (littleEndian)
			bytes.reverse();
		this._setBytes(byteOffset, bytes);
	};
})();

// Stores a signed 8-bit integer (byte) value at the specified byte offset from the start of the view.
DataView.prototype.setInt8 = function (byteOffset, value) {
	if ((value < -128) || (value > 127))
	{
		var e = new TypeError("Int8 value must be between -128 and 127");
		e.description = e.message;
		throw e;
	}
	this._setBytes(byteOffset, [value]);
};

// Stores a signed 16-bit integer (short) value at the specified byte offset from the start of the view.
DataView.prototype.setInt16 = function (byteOffset, value, littleEndian) {
	if ((value < -32768) || (value > 32767))
	{
		var e = new TypeError("Int16 value must be between -32768 and 32767");
		e.description = e.message;
		throw e;
	}
	var bytes = [value>>8, value & 0xFF];
	if (littleEndian)
		bytes.reverse();
	this._setBytes(byteOffset, bytes);
};

// Stores a signed 32-bit integer (long) value at the specified byte offset from the start of the view.
DataView.prototype.setInt32 = function (byteOffset, value, littleEndian) {
	if ((value < -2147483648) || (value > 2147483647))
	{
		var e = new TypeError("Int32 value must be between -2147483648 and 2147483647");
		e.description = e.message;
		throw e;
	}
	var bytes = [value >> 24, value >> 16 & 0xFF, value >> 8 & 0xFF, value & 0xFF];
	if (littleEndian)
		bytes.reverse();
	this._setBytes(byteOffset, bytes);
};

// Stores a signed 64-bit integer (long long) value at the specified byte offset from the start of the view.
// See "Majerus.Automation.Int64" Convert method for details on supported formats.
DataView.prototype.setVarInt64 = function (byteOffset, value, littleEndian) {
	var VarInt64 = new ActiveXObject("Majerus.Automation.Int64");
	var val;
	try {
		val = VarInt64.Convert(value);
	} catch (ex) {
		var e = new TypeError("setVarInt64 value type mismatch")
	}
	var bytes = [];
	for (var i = 0; i < 8; i++)
	{
		// 0..255 values can be coerced into JScript Numbers without
		// any extra step. No need to go through VarInt64.CDec.
		bytes.unshift(Number(VarInt64.BitAnd(val, 0xFF)));
		val = VarInt64.BitRShift(val, 8);
	}
	if (littleEndian)
		bytes.reverse();
	this._setBytes(byteOffset, bytes);
};

// Stores an unsigned 8-bit integer (unsigned byte) value at the specified byte offset from the start of the view.
DataView.prototype.setUint8 = function (byteOffset, value) {
	if ((value < 0) || (value > 255))
	{
		var e = new TypeError("Uint8 value must be between 0 and 255");
		e.description = e.message;
		throw e;
	}
	var bytes = [value];
	this._setBytes(byteOffset, bytes);
};

// Stores an unsigned 16-bit integer (unsigned short) value at the specified byte offset from the start of the view.
DataView.prototype.setUint16 = function (byteOffset, value, littleEndian) {
	if ((value < 0) || (value > 65535))
	{
		var e = new TypeError("Uint16 value must be between 0 and 65535");
		e.description = e.message;
		throw e;
	}
	var bytes = [value>>8, value & 0xFF];
	if (littleEndian)
		bytes.reverse();
	this._setBytes(byteOffset, bytes);
};

// Stores an unsigned 32-bit integer (unsigned long) value at the specified byte offset from the start of the view.
DataView.prototype.setUint32 = function (byteOffset, value, littleEndian) {
	if ((value < 0) || (value > 4294967295))
	{
		var e = new TypeError("Uint32 value must be between 0 and 4294967295");
		e.description = e.message;
		throw e;
	}
	var bytes = [value >> 24, value >> 16 & 0xFF, value >> 8 & 0xFF, value & 0xFF];
	if (littleEndian)
		bytes.reverse();
	this._setBytes(byteOffset, bytes);
};

// Stores an Automation unsigned 64-bit integer (unsigned long long) value at the specified byte offset from the start of the view.
// See "Majerus.Automation.UInt64" Convert method for details on supported formats.
DataView.prototype.setVarUint64 = function (byteOffset, value, littleEndian) {
	var VarUint64 = new ActiveXObject("Majerus.Automation.UInt64");
	var val;
	try {
		val = VarUint64.Convert(value);
	} catch (ex) {
		var e = new TypeError("setVarUint64 value type mismatch")
	}
	var bytes = [];
	for (var i = 0; i < 8; i++)
	{
		// 0..255 values can be coerced into JScript Numbers without
		// any extra step. No need to go through VarUInt64.CDec.
		bytes.unshift(Number(VarUint64.BitAnd(val, 0xFF)));
		val = VarUint64.BitRShift(val, 8);
	}
	if (littleEndian)
		bytes.reverse();
	this._setBytes(byteOffset, bytes);
};

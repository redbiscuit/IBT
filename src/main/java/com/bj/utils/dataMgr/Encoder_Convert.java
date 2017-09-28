package com.bj.utils.dataMgr;


import java.io.UnsupportedEncodingException;

public class Encoder_Convert {
	
	public static String convertChinese(String str) {
		String str1 = "";

		if (str == null || str.length() <= 0) {
			System.out.println("Invalid string!");
		}

		int len = 0;

		char c;
		for (int i = 0; i < str.length(); i++) {
			c = str.charAt(i);

			if (isChinese(c)) {
				str1 += "EN";
				len += 2;
			} else {
				str1 += c;
				len++;
			}
		}

		return str1;
	}

	public static boolean isChinese(char c) {

		Character.UnicodeBlock ub = Character.UnicodeBlock.of(c);

		if (ub == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS

		|| ub == Character.UnicodeBlock.CJK_COMPATIBILITY_IDEOGRAPHS

		|| ub == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_A

		|| ub == Character.UnicodeBlock.GENERAL_PUNCTUATION

		|| ub == Character.UnicodeBlock.CJK_SYMBOLS_AND_PUNCTUATION

		|| ub == Character.UnicodeBlock.HALFWIDTH_AND_FULLWIDTH_FORMS) {

			return true;

		}

		return false;

	}

	/**
	 * UTF-8 to GBK
	 * 
	 * @param text
	 * @return
	 */
	public static String UTF82GBK(String text) {
		String result = "";
		try {
			result = new String(text.getBytes("UTF-8"), "GBK");
		} catch (UnsupportedEncodingException ex) {
			result = ex.toString();
		}
		return result;
	}

	/**
	 * ISO-8859-1 to GB2312
	 * 
	 * @param text
	 * @return
	 */
	public String ISO2GB(String text) {
		String result = "";
		try {
			result = new String(text.getBytes("ISO-8859-1"), "GB2312");
		} catch (UnsupportedEncodingException ex) {
			result = ex.toString();
		}
		return result;
	}
	
	/**
	 * ISO-8859-1 to GB2312
	 * 
	 * @param text
	 * @return
	 */
	public String ISO2UTF8(String text) {
		String result = "";
		try {
			result = new String(text.getBytes("ISO-8859-1"), "UTF-8");
		} catch (UnsupportedEncodingException ex) {
			result = ex.toString();
		}
		return result;
	}

	/**
	 * GB2312 to ISO-8859-1
	 * 
	 * @param text
	 * @return
	 */
	public String GB2ISO(String text) {
		String result = "";
		try {
			result = new String(text.getBytes("GB2312"), "ISO-8859-1");
		} catch (UnsupportedEncodingException ex) {
			ex.printStackTrace();
		}
		return result;
	}

	/**
	 * Utf8URL encode
	 * 
	 * @param
	 * @return
	 */
	public String Utf8URLencode(String text) {
		StringBuffer result = new StringBuffer();
		for (int i = 0; i < text.length(); i++) {
			char c = text.charAt(i);
			if (c >= 0 && c <= 255) {
				result.append(c);
			} else {
				byte[] b = new byte[0];
				try {
					b = Character.toString(c).getBytes("UTF-8");
				} catch (Exception ex) {
				}
				for (int j = 0; j < b.length; j++) {
					int k = b[j];
					if (k < 0)
						k += 256;
					result.append("%" + Integer.toHexString(k).toUpperCase());
				}
			}
		}
		return result.toString();
	}

	/**
	 * Utf8URLdecode
	 * 
	 * @param text
	 * @return
	 */
	public static String Utf8URLdecode(String text) {
		String result = "";
		int p = 0;
		if (text != null && text.length() > 0) {
			text = text.toLowerCase();
			p = text.indexOf("%e");
			if (p == -1)
				return text;
			while (p != -1) {
				result += text.substring(0, p);
				text = text.substring(p, text.length());
				if (text == "" || text.length() < 9)
					return result;
				result += CodeToWord(text.substring(0, 9));
				text = text.substring(9, text.length());
				p = text.indexOf("%e");
			}
		}
		return result + text;
	}

	/**
	 * utf8URL
	 * 
	 * @param text
	 * @return
	 */
	private static String CodeToWord(String text) {
		String result;
		if (Utf8codeCheck(text)) {
			byte[] code = new byte[3];
			code[0] = (byte) (Integer.parseInt(text.substring(1, 3), 16) - 256);
			code[1] = (byte) (Integer.parseInt(text.substring(4, 6), 16) - 256);
			code[2] = (byte) (Integer.parseInt(text.substring(7, 9), 16) - 256);
			try {
				result = new String(code, "UTF-8");
			} catch (UnsupportedEncodingException ex) {
				result = null;
			}
		} else {
			result = text;
		}
		return result;
	}

	/**
	 *check utf8code
	 * 
	 * @param text
	 * @return
	 */
	private static boolean Utf8codeCheck(String text) {
		String sign = "";
		if (text.startsWith("%e"))
			for (int i = 0, p = 0; p != -1; i++) {
				p = text.indexOf("%", p);
				if (p != -1)
					p++;
				sign += p;
			}
		return sign.equals("147-1");
	}

	/**
	 * is uft8 url
	 * 
	 * @param text
	 * @return
	 */
	public static boolean isUtf8Url(String text) {
		text = text.toLowerCase();
		int p = text.indexOf("%");
		if (p != -1 && text.length() - p > 9) {
			text = text.substring(p, p + 9);
		}
		return Utf8codeCheck(text);
	}

	/**
	 * get encoding
	 * 
	 * @param str
	 * @return
	 */
	public static String getEncoding(String str) {
		String encode = "GB2312";
		try {
			if (str.equals(new String(str.getBytes(encode), encode))) {
				String s = encode;
				return s;
			}
		} catch (Exception exception) {
		}
		encode = "ISO-8859-1";
		try {
			if (str.equals(new String(str.getBytes(encode), encode))) {
				String s1 = encode;
				return s1;
			}
		} catch (Exception exception1) {
		}
		encode = "UTF-8";
		try {
			if (str.equals(new String(str.getBytes(encode), encode))) {
				String s2 = encode;
				return s2;
			}
		} catch (Exception exception2) {
		}
		encode = "GBK";
		try {
			if (str.equals(new String(str.getBytes(encode), encode))) {
				String s3 = encode;
				return s3;
			}
		} catch (Exception exception3) {
		}
		return "";
	}

	/**
	 * is GBCode?
	 * 
	 * @param strIn
	 * @return
	 */
	public static boolean isGBCode(String strIn)

	{
		char ch1, ch2;

		if (strIn.length() >= 2) {
			ch1 = strIn.charAt(0);
			ch2 = strIn.charAt(1);

			if (ch1 >= 176 && ch1 <= 247 && ch2 >= 160 && ch2 <= 254)

				return true;

			else
				return false;

		}

		else
			return false;

	}

	/**
	 * isGBKCode
	 * 
	 * @param strIn
	 * @return
	 */

	public static boolean isGBKCode(String strIn)

	{
		char ch1, ch2;

		if (strIn.length() >= 2) {
			ch1 = strIn.charAt(0);
			ch2 = strIn.charAt(1);

			if (ch1 >= 129 && ch1 <= 254 && ch2 >= 64 && ch2 <= 254)

				return true;

			else
				return false;

		}

		else
			return false;

	}
}
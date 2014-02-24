/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * 
 * Copyright 2013 Joseph Yuan
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * 
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */


package excelUtils;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Workbook;

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * 
 * class Helper
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Helper methods and miscellaneous tools for Extension of Apache POI
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * To do:
 * 		[ ] work on date handling
 * 		[ ] work on CSV parsing to further generalize
 * 		[ ] String < - > Date 
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

public class Helper {
	public final static String DATE_REGEX_4_DIGIT_YEAR = "("
			+ "(19|20)[0-9]{2}"
			+ "([-/.\\\\]{1})"
			+ "[0?[1-9]|[1-9]|1[012]]{1,2}"
			+ "\\3"
			+ "([0?[1-9]|[1-9]|1[0-9]|2[0-9]|3[01]]{1,2})"
			+ ")"
			+ "|"
			+ "("
			+ "[0?[1-9]|[1-9]|1[012]]{1,2}"
			+ "([-/.\\\\]{1})"
			+ "([0?[1-9]|[1-9]|1[0-9]|2[0-9]|3[01]]{1,2})"
			+ "\\6"
			+ "(19|20)[0-9]{2}"
			+ ")";
	public final static String DATE_REGEX_2_DIGIT_YEAR = "("
			+ "[0-9]{2}"
			+ "([-/.\\\\]{1})"
			+ "[0?[1-9]|[1-9]|1[012]]{1,2}"
			+ "\\3"
			+ "([0?[1-9]|[1-9]|1[0-9]|2[0-9]|3[01]]{1,2})"
			+ ")"
			+ "|"
			+ "("
			+ "[0?[1-9]|[1-9]|1[012]]{1,2}"
			+ "([-/.\\\\]{1})"
			+ "([0?[1-9]|[1-9]|1[0-9]|2[0-9]|3[01]]{1,2})"
			+ "\\6"
			+ "[0-9]{2}"
			+ ")";
	/* Date Handling tools */
	public static boolean isWeekend(Calendar date) {
		return date.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY ||
				date.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY;
	}
	public static Calendar getToday() {
		return Calendar.getInstance();
	}
	public static Calendar getYesterday() {
		Calendar date = Calendar.getInstance();
		date.add(Calendar.DATE,-1);
		return date;
	}
	public static Calendar getLastWeekday() {
		return getLastWeekday(getYesterday());
	}
	public static Calendar getLastWeekday(Calendar date) {
		Calendar _date = Calendar.getInstance();
		_date.setTime(date.getTime());
		while (isWeekend(_date)) {
			_date.add(Calendar.DATE,-1);
		}
		return _date;
	}
	public static boolean isDate(String s) {
		Pattern pattern = Pattern.compile(DATE_REGEX_4_DIGIT_YEAR);
		String[] terms = s.split(" ");
		Matcher matcher;
		for (String term: terms) {
			matcher = pattern.matcher(term);
			while (matcher.matches()) {
				return true;
			}
		}
		pattern = Pattern.compile(DATE_REGEX_2_DIGIT_YEAR);
		terms = s.split(" ");
		for (String term: terms) {
			matcher = pattern.matcher(term);
			while (matcher.matches()) {
				return true;
			}
		}
		return false;
	}
	public static String parseDate(String entry) {
		Pattern pattern = Pattern.compile(DATE_REGEX_4_DIGIT_YEAR);
		String[] terms = entry.split(" ");
		Matcher matcher;
		for (String term: terms) {
			matcher = pattern.matcher(term);
			while (matcher.matches()) {
				return matcher.group();
			}
		}
		pattern = Pattern.compile(DATE_REGEX_2_DIGIT_YEAR);
		terms = entry.split(" ");
		for (String term: terms) {
			matcher = pattern.matcher(term);
			while (matcher.matches()) {
				return matcher.group();
			}
		}
		return "";
	}
	public static int compareDates(String date1, String date2, String format) {
		/*
		// A Nice Regex Dream, maybe one day
		 * CHANGE TO THE FILE
		System.out.println("ORIGINAL: " + format);
		String format_regex = format.replaceAll("(yy(yy)?)", "(?<year>$1)");
		format_regex = format_regex.replaceAll("(mm(mm)?)","(?<month>$1)");
		format_regex = format_regex.replaceAll("(d(d)?)", "(?<day>$1)");
		format_regex = format_regex.replaceAll("(?<!\\<)d","\\\\\\\\d");
		TESTING ON MAC SIDE  LOUD AND CLEAR
		TESTING ON PC SIDE   OVER AND OUT AND CHANGE
		format_regex = format_regex.replaceAll("(?<!\\<)m","\\\\\\\\d");
		format_regex = format_regex.replaceAll("(?<!\\<)y","\\\\\\\\d");
		System.out.println("FINAL: " + format_regex);
		Pattern p = Pattern.compile(format_regex);
		Matcher m = p.matcher(date1);
		while (m.find()){
			System.out.println(m.group());
		}*/
		if (format.length() == 0) {
			return 0;
		}
		char c;
		int m_start = 0,m_count = 0,
				d_start = 0,d_count = 0,
				y_start = 0,y_count = 0;
		for (int i = 0; i < format.length(); i++) {
			c = format.charAt(i);
			if (c == 'm') { 
				if (m_count == 0) {
					m_start = i;
				}
				m_count++;
			}
			if (c == 'd') { 
				if (d_count == 0) {
					d_start = i;
				}
				d_count++;
			}
			if (c == 'y') { 
				if (y_count == 0) {
					y_start = i;
				}
				y_count++;
			}
		}
		if (y_count > 0) {
			if (Integer.parseInt(date1.substring(y_start, y_start+y_count)) > Integer.parseInt(date2.substring(y_start,y_start+y_count))) {
				return 1;
			} else if (Integer.parseInt(date1.substring(y_start, y_start+y_count)) < Integer.parseInt(date2.substring(y_start,y_start+y_count))) {
				return -1;
			}
		}
		if (m_count > 0) {
			if (Integer.parseInt(date1.substring(m_start, m_start+m_count)) > Integer.parseInt(date2.substring(m_start,m_start+m_count))) {
				return 1;
			} else if (Integer.parseInt(date1.substring(m_start, m_start+m_count)) < Integer.parseInt(date2.substring(m_start,m_start+m_count))) {
				return -1;
			}
		}
		if (d_count > 0) {
			if (Integer.parseInt(date1.substring(d_start, d_start+d_count)) > Integer.parseInt(date2.substring(d_start,d_start+d_count))) {
				return 1;
			} else if (Integer.parseInt(date1.substring(d_start, d_start+d_count)) < Integer.parseInt(date2.substring(d_start,d_start+d_count))) {
				return -1;
			}
		}
		return 0;
	}
	public static Calendar stringToCalendar(String date) {
		return stringToCalendar(date,"mm/dd/yy");
	}
	public static Calendar stringToCalendar(String date, String pattern) {
		return null;
	}
	public static Date stringToDate(String date) {
		return stringToDate(date,"mm/dd/yy");
	}
	public static Date stringToDate(String date, String pattern) {
		return null;
	}
	/* Parsers */
	public static ArrayList<String> parseCSVLine(String CSVLine,char delimChar,char quotChar) {
		char itr;
		boolean inQuotedValue = false;
		String buffer = "";
		ArrayList<String> parsedLine = new ArrayList<String>();
		for (int i = 0; i < CSVLine.length(); i++) {
			itr = CSVLine.charAt(i);
			if (itr == delimChar) {
				if (!inQuotedValue) {
					parsedLine.add(buffer);
					buffer = "";
				}
			} else if (itr == quotChar) {
				inQuotedValue = !inQuotedValue;
			} else {
				buffer += itr;
			}
		}
		if (buffer.length() > 0) {
			parsedLine.add(buffer);
		}
		String entry;
		for (int i = 0; i < parsedLine.size(); i++) {
			entry = parsedLine.get(i);
			entry.trim();
			if (entry.startsWith("(") && entry.endsWith(")") && isNumeric(entry.substring(1,entry.length()-1))) {
				parsedLine.set(i, "-" + entry.substring(1,entry.length()-1)); 
			}
		}
		return parsedLine;
	}
	public static class HTMLNode {
		HTMLNode firstChild;
		HTMLNode firstNeighbor;
		String tag;
		String attr;
		String content;
		public HTMLNode() {}
		public HTMLNode(HTMLNode firstChild, HTMLNode firstNeighbor, String tag, String attr, String content) {
			this.tag = tag;
			this.attr = attr;
			this.firstChild = firstChild;
			this.firstNeighbor = firstNeighbor;
			this.content = content;
		} 
		@Override
		public String toString() {
			return "<" + tag + ">";
		}
	}
	private static int prevIndex(String str, char c) {
		return prevIndex(str,c,str.length()-1);
	}
	private static int prevIndex(String str, char c, int fromIndex) {
		for (int i = fromIndex; i >= 0; i--) {
			if (str.charAt(i) == c) {
				return i;
			}
		}
		return -1;
	}
	private static int max(int a, int b) {
		if (a > b) {
			return a;
		} else {
			return b;
		}
	}
	private static int min(int a, int b) {
		if (a < b) {
			return a;
		} else {
			return b;
		}
	}
	private static int minPos(int a, int b) {
		if (a >= 0 && b >= 0) {
			return min(a,b);
		} else if (a >= 0 && b < 0) {
			return a;
		} else if (b >= 0 && a < 0) {
			return b;
		} else {
			return Integer.MAX_VALUE;
		}
	}
	private static String safeSubstring(String str, int i, int f) {
		try {
			i = max(0,i);
			i = min(i,str.length());
			// Limit to 0
			f = max(0,f);
			f = min(f,str.length());
			return str.substring(i, f);
		} catch (Exception e) {
			return "";
		}
	}
	public static HTMLNode parseHTML(String html) {
		html = html.trim();
		if (html.length() == 0) return null;
		char c;
		String tagName = "", childContent = "", attr = "", lastTag = "", content = "";
		int n, lastIndex = 0, remainderIndex = -1;
		boolean isSingleton = false;
		for (int i = 0, L = html.length(), depth = 0; i < L && i >= 0; lastIndex = i, i = html.indexOf('<',i+1)) {
			c = html.charAt(i);
			if (i == 0) {
				if (c != '<') {
					childContent = safeSubstring(html, 0, (remainderIndex = minPos(html.indexOf('<'),L)) );
					tagName = "text";
					break;
				} else {
					if (i < L-1 && html.charAt(i+1) == '!') {
						// Comment
						if (safeSubstring(html,i+2,(n=i+4)).equals("--")) {
							// <!-- ...comment... -->
							tagName = "comment";
							childContent = safeSubstring(html,n,(remainderIndex = html.indexOf("-->",n)));
							remainderIndex += "-->".length();
							break;
						} else {
							// <!DOCTYPE html>
							childContent = safeSubstring(html,i,(remainderIndex = html.indexOf('>',i)+1));
						}
					}
					tagName = safeSubstring(html,i+1,(n = minPos(html.indexOf('>', i+1),html.indexOf(' ',i+1)))).toLowerCase();
					attr = safeSubstring(html,n,(n = html.indexOf('>',n)));
					i = n+1;
					// Check for singleton
					isSingleton = tagName.equals("meta") || tagName.equals("br") ||
							tagName.equals("hr") || tagName.equals("link") || tagName.equals("area") ||
							tagName.equals("base") || tagName.equals("col") || tagName.equals("command") ||
							tagName.equals("input") || tagName.equals("embed") || tagName.equals("img") ||
							tagName.equals("param")  || tagName.equals("source");
					if (isSingleton) {
						remainderIndex = html.indexOf('>') + 1;
						break;
					}
					depth++;
				}
			} else {
				if (lastIndex > 0) {
					childContent += safeSubstring(html,lastIndex,i);
				}
				if (c == '<') {
					if (i < L-1 && html.charAt(i+1) == '/') {
						lastTag = safeSubstring(html,i+2,(n = html.indexOf('>', i+2))).toLowerCase();
						if (lastTag.equalsIgnoreCase(tagName)) {
							depth--;
						}
						if (depth <= 0) {
							remainderIndex = n+1;
							if (lastIndex > 0) {
								childContent += safeSubstring(html,lastIndex,i);
							}
							break;
						}
					} else {
						if (safeSubstring(html,i+1,i+4).equals("!--")) { // Comment
							lastTag = "comment";
							n = html.indexOf("-->",i+4) + "-->".length();
						} else {
							lastTag = safeSubstring(html,i+1,(n = min(html.indexOf('>', i+1),html.indexOf(' ',i+1)))).toLowerCase();
						}
						if (lastTag.equalsIgnoreCase(tagName)) {
							depth++;
						}
						
					}
				}
			}
		}
		String remainder = safeSubstring(html,remainderIndex,html.length());
		content = childContent;
		if (isSingleton || tagName.equalsIgnoreCase("comment") || tagName.equalsIgnoreCase("text")) {
			childContent = ""; // Have no children
		}
		System.out.println();
		return new HTMLNode(parseHTML(childContent),parseHTML(remainder),tagName,attr,content);
	}
	public static void printHTML(HTMLNode n) {
		printHTML(n, 0);
	}
	public static void printHTML(HTMLNode n, int level) {
		if (n != null) {
			String prefix = "";
			for (int i = 0; i < level; i++) {
				prefix += '\t';
			}
			System.out.print(prefix + "<" + n.tag + n.attr + ">");
			if (n.firstChild != null) {
				System.out.println();
			}
			if (n.firstChild != null) {
				System.out.print('t'+prefix);
			}
			System.out.print(n.content);
			printHTML(n.firstChild,level+1);
			if (n.firstChild != null) {
				System.out.print(prefix);
			}
			System.out.println("</" + n.tag + ">");
			printHTML(n.firstNeighbor,level);

		}
	}
	public static boolean isNumeric(String str) {
		String s = str;
		if (s.startsWith("-")) { s = s.substring(1);}
		char c;
		int i,L = s.length();
		boolean decimalHit = false;
		for (i = 0; i < L; i++) {
			c = s.charAt(i);
			if (c < '0' || c > '9') {
				if (c == '.' && !decimalHit) {
					decimalHit = true;
					continue;
				}
				return false;
			}
		}
		return true;
	}
	public static boolean isUpperAlpha(char c) {
		return 'A' <= c && 'Z' >= c;
	}
	public static boolean isLowerAlpha(char c) {
		return 'a' <= c && 'z' >= c;
	}
	public static boolean isAlpha(char c) {
		return isUpperAlpha(c) || isLowerAlpha(c);
	}
	public static boolean isNumeric(char c) {
		return '0' <= c && '9' >= c;
	}
	public static String formatYesterday() {
		return formatDate("MM.dd.yy",0,0,-1);
	}
	public static String formatYesterday(String format) {
		return formatDate(format,0,0,-1);
	}
	public static String formatToday() {
		return formatDate("MM.dd.yy",0,0,0);
	}
	public static String formatToday(String format) {
		return formatDate(format,0,0,0);
	}
	public static String formatLastWeekday() {
		return formatDate("MM.dd.yy",Helper.getLastWeekday());
	}
	public static String formatLastWeekday(String format) {
		return formatDate(format,Helper.getLastWeekday());
	}
	public static String formatDate(String format, int offset_years, int offset_months, int offset_days) {
		Calendar date = Calendar.getInstance();
		date.add(Calendar.YEAR, offset_years);
		date.add(Calendar.MONTH, offset_months);
		date.add(Calendar.DATE, offset_days);
		return formatDate(format,date);
	}
	public static String formatDate(Calendar c) {
		return (new SimpleDateFormat("MM.dd.yy")).format(c.getTime());
	}
	public static String formatDate(String format, Calendar c) {
		return (new SimpleDateFormat(format)).format(c.getTime());
	} 
	public static String parseToRegExcel(String raw) {
		char c;
		String buf = "", regExcel = "";
		for (int i = 0; i < raw.length(); i++) {
			c = raw.charAt(i);
			if (isNumeric(c) && i+1 < raw.length() && isNumeric(raw.charAt(i+1))) {
				while (isNumeric(c)) {
					buf += c;
					i++;
					c = raw.charAt(i);
				}
				regExcel += parseRegExcelNumber(buf)+c;
				buf = "";
			} else {
				regExcel += c;
			}
		}
		return regExcel;
	}
	private static String parseRegExcelNumber(String number) {
		String today_d = formatToday("dd"), 
				today_m = formatToday("MM"), 
				today_y = formatToday("yyyy");
		String regex = "(" + today_y.substring(0,2) + ")?" + today_y.substring(2) + today_m + today_d;
		Pattern p = Pattern.compile(regex);
		Matcher m = p.matcher(number);
		String buf = "";
		if (m.find()) {
			if (m.group().length() == 6) {
				buf = (m.start() == 0 ? "" : "<NUMBER>" ) + "<DATE:TODAY:YYMMDD>" + (m.end() == number.length() ? "" : "<NUMBER>");
			} else {
				buf = (m.start() == 0 ? "" : "<NUMBER>" ) + "<DATE:TODAY:YYYYMMDD>" + (m.end() == number.length() ? "" : "<NUMBER>");
			}
			return buf;
		}
		String yesterday_d = formatYesterday("dd"), 
				yesterday_m = formatYesterday("MM"), 
				yesterday_y = formatYesterday("yyyy");
		regex = "(" + yesterday_y.substring(0,2) + ")?" + yesterday_y.substring(2) + yesterday_m + yesterday_d;
		p = Pattern.compile(regex);
		m = p.matcher(number);
		buf = "";
		if (m.find()) {
			if (m.group().length() == 6) {
				buf = (m.start() == 0 ? "" : "<NUMBER>" ) + "<DATE:YESTERDAY:YYMMDD>" + (m.end() == number.length() ? "" : "<NUMBER>");
			} else {
				buf = (m.start() == 0 ? "" : "<NUMBER>" ) + "<DATE:YESTERDAY:YYYYMMDD>" + (m.end() == number.length() ? "" : "<NUMBER>");
			}
			return buf;
		}
		return "<NUMBER>";

	}
	public static String parseToRegex(String raw) {
		char c;
		int depth = 0;
		String buf = "",regex = "";
		for (int i = 0; i < raw.length() ;i++) {
			c = raw.charAt(i);
			if (i > 0 && c == '.' && raw.charAt(i-1) != '\\') {
				buf = "\\.";
			} else if (c == '<') {
				depth++;
				buf = "";
				while (depth > 0 && i < raw.length()) {
					i++;
					c = raw.charAt(i);
					if (c == '<') {
						depth++;
						buf += c;
					} else if (c == '>') {
						depth--;
						if (depth == 0) {
							break;
						}
					} else {
						buf += c;
					}
				}
				if (buf.equalsIgnoreCase("NUMBER")) {
					buf = "\\d+";
				} else if (buf.toLowerCase().contains("DATE".toLowerCase())) {
					buf = "{$@"+buf+"@$}";
				} else if (buf.equalsIgnoreCase("RANDOM")) {
					buf = ".*?";
				} else {
					buf = ".*";
				}
			} else {
				buf += c;
			}
			regex += buf;
			buf = "";
		}
		int begin, end;
		String format;
		String ref_regex = regex.toUpperCase(), ref_format;
		begin = ref_regex.indexOf("{$@DATE:") + ("{$@DATE:").length();
		end = ref_regex.indexOf("@$}",begin);
		while (begin > 0 && end > 0) {
			format = regex.substring(begin,end);
			ref_format = format.toUpperCase();
			if (ref_format.startsWith("TODAY:")) {
				begin += "TODAY:".length();
				format = regex.substring(begin,end);
				format = format.replace('m', 'M').replace('D', 'd').replace('Y', 'y');
				regex = regex.substring(0,begin-("{$@DATE:TODAY:").length()) + formatToday(format) + regex.substring(end+("@$}").length());
			} else if (ref_format.contains("YESTERDAY")) {
				begin += ("YESTERDAY:").length();
				format = regex.substring(begin,end);
				format = format.replace('m', 'M').replace('D', 'd').replace('Y', 'y');
				regex = regex.substring(0,begin-("{$@DATE:YESTERDAY:").length()) + formatYesterday(format) + regex.substring(end+("@$}").length());
			} else if (ref_format.contains("LASTWEEKDAY")) {
				begin += ("LASTWEEKDAY:").length();
				format = regex.substring(begin,end);
				format = format.replace('m', 'M').replace('D', 'd').replace('Y', 'y');
				regex = regex.substring(0,begin-("{$@DATE:LASTWEEKDAY:").length()) + formatLastWeekday(format) + regex.substring(end+("@$}").length());
			} else {
				format = format.replace('m', 'M').replace('D', 'd').replace('Y', 'y');
				regex = regex.substring(0,begin-("{$@DATE:".length())) + formatToday(format) + regex.substring(end+("@$}").length());
			}
			begin = regex.indexOf("{$@DATE:")+("{$@DATE:").length();
			end = regex.indexOf("@$}",begin);
		}
		return regex;
	}
	public static void main(String args[]) {
		try { //C:\Users\Joe\Desktop\Projects\Rec
			Workbook wb = ExcelUtils.openWorkbook(FileUtils.joinPath("C:","Users","Joe","Desktop","Projects","Rec","GS-SDI-Account_Balances_By_Currency-08_36_GMT-20140219-17040-0.html"));
			System.out.println(ExcelUtils.saveWorkbook(wb, "HTMLtest.xls", FileUtils.joinPath("C:","Users","Joe","Desktop","Projects","Rec"), false, true, true));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}

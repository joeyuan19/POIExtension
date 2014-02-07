package excelUtils;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
	public final static String DATE_REGEX = "("
			+ "(19|20)?[0-9]{2}"
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
			+ "(19|20)?[0-9]{2}"
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
		Pattern pattern = Pattern.compile(DATE_REGEX);
		String[] terms = s.split(" ");
		Matcher matcher;
		for (String term: terms) {
			matcher = pattern.matcher(term);
			while (matcher.find()) {
				return true;
			}
		}
		return false;
	}
	public static String parseDate(String entry) {
		Pattern pattern = Pattern.compile(DATE_REGEX);
		String[] terms = entry.split(" ");
		Matcher matcher;
		for (String term: terms) {
			matcher = pattern.matcher(term);
			while (matcher.find()) {
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
		TESTING ON PC SIDE
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
	/**/
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
	public static boolean isNumeric(String s) {
		try {
			Float.parseFloat(s);
			return true;
		} catch (NumberFormatException e) {
			try {
				Integer.parseInt(s);
				return true;
			} catch (NumberFormatException e1) {
				try {
					Long.parseLong(s);
					return true;
				} catch (NumberFormatException e2) {
					return false;
				}
			}
		}
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
	public static String formatLastWeekdayFormatted() {
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
			System.out.println(number+ " " + m.group() + " " + m.start());
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
		System.out.println(regex);
		if (m.find()) {
			System.out.println(number+ " " + m.group() + " " + m.start());
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
		begin = regex.indexOf("{$@DATE:")+("{$@DATE:").length();
		end = regex.indexOf("@$}",begin);
		while (begin > 0 && end > 0) {
			format = regex.substring(begin,end);
			if (format.toLowerCase().contains("TODAY".toLowerCase())) {
				begin = regex.toLowerCase().indexOf("{$@DATE:TODAY:".toLowerCase())+("{$@DATE:TODAY:").length();
				format = regex.substring(begin,end);
				format = format.replace('m', 'M').replace('D', 'd').replace('Y', 'y');
				regex = regex.substring(0,begin-("{$@DATE:TODAY:").length()) + formatToday(format) + regex.substring(end+("@$}").length());
			} else if (format.toLowerCase().contains("TODAY".toLowerCase())) {
				begin = regex.toLowerCase().indexOf("{$@DATE:YESTERDAY:".toLowerCase())+("{$@DATE:YESTERDAY:").length();
				format = regex.substring(begin,end);
				format = format.replace('m', 'M').replace('D', 'd').replace('Y', 'y');
				regex = regex.substring(0,begin-("{$@DATE:YESTERDAY:").length()) + formatYesterday(format) + regex.substring(end+("@$}").length());
			} else {
				format = format.replace('m', 'M').replace('D', 'd').replace('Y', 'y');
				regex = regex.substring(0,begin-("{$@DATE:").length()) + formatToday(format) + regex.substring(end+("@$}").length());
			}
			begin = regex.indexOf("{$@DATE:")+("{$@DATE:").length();
			end = regex.indexOf("@$}",begin);
		}
		return regex;
	}
	public static void main(String args[]) {
		System.out.println(FileUtils.parseExt(parseToRegExcel("bpmon.csv.140203110802.csv")));
	}
}

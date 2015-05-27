package resources;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class GeneralUtils {

	public static final Pattern tag = Pattern.compile("<[^>]*>");

	public static String formatHTML(String HTML) {
		HTML = HTML.replaceAll("[\n\r]", "");
		HTML = HTML.replaceAll(" {2,}", "");
		StringBuilder string = new StringBuilder();
		Matcher m = tag.matcher(HTML);

		int tabs = 0;
		int startIndex = 0;
		int endIndex = 0;
		while (m.find()) {

			startIndex = m.start();
			if (startIndex > endIndex) {
				for (int i = 0; i < tabs; i++) {
					string.append("\t");
				}
				string.append(HTML.substring(endIndex, startIndex))
						.append('\n');
			}

			if (m.group().startsWith("</")) {
				tabs = tabs - 1;
			}

			// System.out.println(tabs + " (" + m.group().startsWith("</") +
			// "): "
			// + m.group());

			for (int i = 0; i < tabs; i++) {
				string.append("\t");
			}

			string.append(m.group());
			string.append('\n');

			if (m.group().startsWith("<") && !m.group().endsWith("/>")
					&& !m.group().startsWith("<!--")
					&& !m.group().startsWith("<meta")
					&& !m.group().startsWith("</")
					&& !m.group().startsWith("<br")
					&& !m.group().startsWith("<hr")) {
				tabs = tabs + 1;
			}
			endIndex = m.end();
		}
		return string.toString();
	}

	// public static void main(String[] args) {
	// System.out.println(formatHTML("<p>\n" + "    afdbsdgb\n" + "   <br>\n"
	// + "</p>\n"));
	// }

	public static String nameCase(String name) {
		String[] parts = name.split(" ");
		name = "";
		for (String part : parts) {
			name += Character.toUpperCase(part.charAt(0))
					+ part.substring(1).toLowerCase();
			name += " ";
		}
		return name.trim();
	}
}

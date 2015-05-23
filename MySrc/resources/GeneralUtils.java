package resources;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class GeneralUtils {

	public static final Pattern tag = Pattern.compile("<[^>]*>");

	public static String printHTML(String HTML) {
		StringBuilder string = new StringBuilder();
		Matcher m = tag.matcher(HTML);

		int tabs = 0;
		int startIndex = 0;
		int endIndex = 0;
		while (m.find()) {

			startIndex = m.start();
			if (startIndex > endIndex) {
				for (int i = 0; i < tabs; i++) {
					string.append("    ");
				}
				string.append(HTML.substring(endIndex, startIndex))
						.append('\n');
			}

			if (m.group().startsWith("</"))
				tabs = tabs - 1;

			// System.out.println(tabs + " (" + m.group().startsWith("</") +
			// "): "
			// + m.group());

			for (int i = 0; i < tabs; i++) {
				string.append("    ");
			}

			string.append(m.group());
			string.append('\n');

			if (!m.group().endsWith("/>") && !m.group().startsWith("<!--")
					&& !m.group().startsWith("<meta")
					&& !m.group().startsWith("</") && !m.group().equals("<br>")
					&& !m.group().startsWith("<hr"))
				tabs = tabs + 1;
			endIndex = m.end();
		}
		return string.toString();
	}
}

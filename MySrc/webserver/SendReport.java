package webserver;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

import org.joda.time.DateTime;
import org.joda.time.MutableDateTime;

import resources.GeneralUtils;
import sql.SQLUtilities;
import email.EmailMessageCreator;

public class SendReport extends Thread {

	private MutableDateTime whenToSend;
	private boolean running;
	private EmailAddress address;
	private int timesRemaining;

	/**
	 * What day of the month to send the report.
	 * 
	 * @param dayOfMonth
	 */
	public SendReport() {
		// address = new EmailAddress("Molly Underwood",
		// "Molly.Underwood@mechdyne.com");
		address = new EmailAddress("James Talbert",
				"James.Talbert@mechdyne.com");
		timesRemaining = 0;
		whenToSend = null;
	}

	public void usersNotified(DateTime when) {
		whenToSend = when.toMutableDateTime();
		whenToSend.addDays(5);
		timesRemaining = 2;
		start();
	}

	public void usersNotified() {
		usersNotified(DateTime.now());
	}

	@Override
	public void start() {
		if (whenToSend != null) {
			super.start();
		}
	}

	@Override
	public void run() {
		running = true;
		while (running && timesRemaining > 0) {
			try {
				Thread.sleep(3600000);
			} catch (InterruptedException e1) {
				e1.printStackTrace();
			}
			if (whenToSend.isBefore(DateTime.now().getMillis())) {
				System.out.println(whenToSend);
				try {
					sendReport();

					whenToSend.addDays(5);
					timesRemaining--;

				} catch (SQLException e) {
					e.printStackTrace();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
		whenToSend = null;
	}

	private void sendReport() throws Exception {
		List<Object[]> rows = getResponsesAndOverages();
		sort(rows);

		EmailMessageCreator factory = new EmailMessageCreator();
		EmailMessage email = factory.newEmail();
		email.setSubject("Overage response report");
		email.setFrom(new EmailAddress("Overage Notifier",
				"Overage_Notifier@mechdyne.com"));
		String table = getTable(rows);
		MessageBody body = new MessageBody(
				"Here is the list of users with data overages<br>" + "<br>"
						+ table);
		email.setBody(body);
		email.getToRecipients().add(address);
		email.send();
	}

	private static void sort(List<Object[]> rows) {
		Collections.sort(rows, new Comparator<Object[]>() {

			@Override
			public int compare(Object[] o1, Object[] o2) {
				return ((String) o1[0]).compareTo((String) o2[0]);
			}
		});
	}

	private String getTable(Iterable<Object[]> rows) {
		StringBuilder builder = new StringBuilder();
		builder.append("<table border=\"1\">");
		// @FormatterOff
		builder.append("<tr><th>Name</th>" + "<th>Service Number</th>"
					 + "<th>Overage charge ($)</th>"
					 + "<th>Other charges ($)</th>"
					 + "<th>Where to deduct from</th></tr>");
		//@FormatterOn
		for (Object[] row : rows) {
			builder.append("<tr>");
			for (Object cell : row) {
				builder.append("<td>");
				builder.append(GeneralUtils.nameCase(cell.toString()));
				builder.append("</td>");
			}
			builder.append("</tr>");
		}
		builder.append("</table>");
		String table = builder.toString();
		return table;
	}

	private List<Object[]> getResponsesAndOverages() throws SQLException {
		Statement statement = SQLUtilities.getNewStatement();
		ResultSet responses = statement
				.executeQuery("SELECT * FROM responses;");

		Map<Long, String> responsesMap = new HashMap<>();
		while (responses.next()) {
			responsesMap.put(responses.getLong("number"),
					responses.getString("response"));
		}

		List<Object[]> rows = new ArrayList<>();
		ResultSet overages = statement
				.executeQuery("SELECT * FROM internal_data WHERE data_overage>0 OR other_personal_charges>0;");
		while (overages.next()) {
			Long number = overages.getLong("number");
			String name = overages.getString("name");
			Double overage = overages.getDouble("overage_charge");
			Double otherPersonal = overages.getDouble("other_personal_charges");
			String response = responsesMap.get(number);
			if (response == null) {
				response = " ";
			}
			Object[] row = new Object[5];
			row[0] = name;
			row[1] = number;
			row[2] = overage;
			row[3] = otherPersonal;
			row[4] = response;
			rows.add(row);
		}
		return rows;
	}

	public static void main(String[] args) {
		MutableDateTime now = DateTime.now().toMutableDateTime();
		now.setMinuteOfDay(now.getMinuteOfDay() + 1);
		SendReport reporter = new SendReport();
		reporter.usersNotified(DateTime.now());
		reporter.start();
	}
}

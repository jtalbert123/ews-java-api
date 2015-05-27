package webserver;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.ServerSocket;
import java.net.Socket;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JOptionPane;

import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;

import org.joda.time.DateTime;

import resources.GeneralUtils;
import responseListener.BasicResponseDetector;
import sql.ReadDB;
import sql.SQLUtilities;
import email.OverageNotifications;

public class SimpleWebServer implements Runnable {
	private ServerSocket socket;
	public final Thread thread;
	private boolean running;
	private SendReport reporter;

	public SimpleWebServer() throws Exception {

		socket = new ServerSocket(65530);
		thread = new Thread(this);
		BasicResponseDetector listener = new BasicResponseDetector();
		reporter = new SendReport();
	}

	private static Pattern PostRequest = Pattern
			.compile("POST /(.*) HTTP/\\d\\.\\d");

	public void run() {
		running = true;
		while (running) {
			try {
				Socket connection = socket.accept();
				BufferedReader input = new BufferedReader(
						new InputStreamReader(connection.getInputStream()));

				String line = input.readLine();

				// System.out.println(line);

				if (line == null) {
					continue;
				}
				if (line.matches(PostRequest.toString())) {
					Matcher getMessage = PostRequest.matcher(line);
					if (getMessage.find()) {
						String request = getMessage.group(1);
						System.out.println(request);
						if (request.equals("notify")) {
							System.out.println("notifing");

							respond(connection, "notifing");

							sendMessages();
							reporter.usersNotified(DateTime.now());
						} else if (request.equals("clear")) {
							SQLUtilities.clearTable("responses");
							respond(connection, "cleared");
						} else {
							respond(connection, "unrecognized POST request");
						}
						continue;
					}
				}
				respond(connection, "only POST requests accepted");
			} catch (IOException e) {
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	private void sendMessages() throws Exception {
		System.out.println("Sending messages");
		Set<Map<String, Object>> set = ReadDB.getMessageProperties();
		if (set.size() == 0) {
			JOptionPane
					.showMessageDialog(
							null,
							"No entries were found on the server. Please re-generate the report first.\n"
									+ "If the report generation is interrupted, the old data may have been cleared, but not replaced.",
							"Data not Found", JOptionPane.ERROR_MESSAGE);
		}
		OverageNotifications factory = new OverageNotifications(false);
		EmailMessage email = null;
		for (Map<String, Object> properties : set) {
			email = factory.newEmail(properties);
			for (EmailAddress addr : email.getToRecipients())
				System.out.print(addr.toString());
			System.out.println();

			System.out.println(GeneralUtils.formatHTML(email.getBody()
					.toString()));

			email.getToRecipients().clear();
			email.send();
		}
		factory.close();
	}

	private void respond(Socket connection, String response) throws IOException {
		OutputStream out = connection.getOutputStream();
		String statusLine = "HTTP/1.1 200 OK\r\n";
		out.write(statusLine.getBytes("ASCII"));

		String contentLength = "Content-Length: " + response.length() + "\r\n";
		out.write(contentLength.getBytes("ASCII"));

		// signal end of headers
		out.write("\r\n".getBytes("ASCII"));

		// write actual response and flush
		out.write(response.getBytes());
		out.flush();
	}

	public void start() {
		thread.start();
	}

	public void stop() {
		running = false;
	}

	public static void main(String[] args) throws Exception {
		SimpleWebServer sws = new SimpleWebServer();
		sws.start();
		System.out.println("ready, probably");
	}
}

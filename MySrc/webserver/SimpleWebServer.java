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

import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import sql.ReadDB;
import email.OverageNotifications;

public class SimpleWebServer implements Runnable {
	private ServerSocket socket;
	public final Thread thread;

	public SimpleWebServer() throws IOException {
		socket = new ServerSocket(65530);
		thread = new Thread(this);
		// thread.start();
	}

	private static Pattern PostRequest = Pattern
			.compile("POST /(.*) HTTP/\\d\\.\\d");

	public void run() {
		while (true) {
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
							respond(connection, "notifing");
							sendMessages();
						}
						continue;
					}
				}
				respond(connection, "unrecognized data");
			} catch (IOException e) {
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	private void sendMessages() throws Exception {
		Set<Map<String, Object>> set = ReadDB.getMessageProperties();
		OverageNotifications factory = new OverageNotifications();
		for (Map<String, Object> properties : set) {
			EmailMessage email = factory.newEmail(properties);
			for (EmailAddress addr : email.getToRecipients())
				System.out.print(addr.toString());
			System.out.println(';');
		}
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

	public static void main(String[] args) throws IOException {
		SimpleWebServer sws = new SimpleWebServer();
		sws.start();
		System.out.println("ready, probably");
	}
}

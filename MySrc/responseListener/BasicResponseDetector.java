package responseListener;

import static resources.EWSSetup.ShortCallingService;
import static resources.EWSSetup.SubscriptionService;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.enumeration.EventType;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.notification.ItemEvent;
import microsoft.exchange.webservices.data.notification.NotificationEvent;
import microsoft.exchange.webservices.data.notification.NotificationEventArgs;
import microsoft.exchange.webservices.data.notification.StreamingSubscription;
import microsoft.exchange.webservices.data.notification.StreamingSubscriptionConnection;
import microsoft.exchange.webservices.data.notification.StreamingSubscriptionConnection.INotificationEventDelegate;
import microsoft.exchange.webservices.data.notification.StreamingSubscriptionConnection.ISubscriptionErrorDelegate;
import microsoft.exchange.webservices.data.notification.SubscriptionErrorEventArgs;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import sql.SQLUtilities;

public class BasicResponseDetector implements INotificationEventDelegate,
		ISubscriptionErrorDelegate {

	private StreamingSubscriptionConnection connection;

	public BasicResponseDetector() throws Exception {
		ArrayList<FolderId> folders = new ArrayList<>();
		folders.add(Folder.bind(ShortCallingService, WellKnownFolderName.Inbox)
				.getId());
		StreamingSubscription subscription = SubscriptionService
				.subscribeToStreamingNotifications(folders, EventType.Created);
		ArrayList<StreamingSubscription> subscriptions = new ArrayList<>();
		subscriptions.add(subscription);
		connection = new StreamingSubscriptionConnection(SubscriptionService,
				subscriptions, 20);

		connection.addOnNotificationEvent(this);
		connection.addOnDisconnect(this);
		connection.open();

		System.out.println("Ready");
	}

	@Override
	public void notificationEventDelegate(Object sender,
			NotificationEventArgs args) {
		for (NotificationEvent e : args.getEvents()) {
			if (e instanceof ItemEvent) {
				ItemEvent event = (ItemEvent) e;
				try {
					Item item = Item.bind(ShortCallingService,
							event.getItemId());
					if (item instanceof EmailMessage) {
						EmailMessage message = (EmailMessage) item;
						System.out.println("New email recieved");

						if (message.getBody().toString()
								.contains("Sent by Overage Notifier")) {
							System.out
									.println("It's an overage notifiation message");

							handleResponse(message);
							break;
						}
					}
				} catch (Exception e1) {
					e1.printStackTrace();
				}
			}
		}
	}

	@Override
	public void subscriptionErrorDelegate(Object sender,
			SubscriptionErrorEventArgs args) {
		try {
			connection.open();
		} catch (ServiceLocalException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static final Pattern deduction = Pattern.compile(
			"[.\n]*Deduct from ([^<\\.]+)[.\n]*", Pattern.CASE_INSENSITIVE);
	private static final Pattern number = Pattern.compile(
			"Sent for phone number: (\\d{10})", Pattern.CASE_INSENSITIVE);

	private void handleResponse(EmailMessage message)
			throws ServiceLocalException, SQLException {
		String body = message.getBody().toString();
		Matcher m = number.matcher(body);
		if (!m.find())
			return;
		String number = m.group(1);
		Statement statement = SQLUtilities.getNewStatement();
		String sql = "SELECT name, account FROM phone_users WHERE number="
				+ number;
		// System.out.println(sql);
		ResultSet result = statement.executeQuery(sql);
		String name = "";
		long account = 0l;
		if (result.next()) {
			name = result.getString("name");
			account = result.getLong("account");
			System.out.println(name);
			System.out.println(account);
		} else {
			System.out.println("error: line 121");
			return;
		}

		String choice = "";
		boolean found = false;
		m = deduction.matcher(body);
		if (m.find()) {
			System.out.println("correctly formatted response");
			choice = m.group(1).toLowerCase();
			System.out.println(choice);
			found = true;
			sql = String
					.format("INSERT INTO responses (name, number, account, response) "
							+ "VALUES ('%s', %s, %d, '%s') ON DUPLICATE KEY UPDATE response=VALUES(response)",
							name, number, account, choice);
			statement.executeUpdate(sql);
			return;
		}
	}

	public static void main(String[] args) throws Exception {
		Class.forName("sql.SQLUtilities");
		BasicResponseDetector listener = new BasicResponseDetector();
	}
}

package subscription;

import static resources.EWSSetup.ShortCallingService;
import static resources.EWSSetup.SubscriptionService;

import java.util.ArrayList;

import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
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
import resources.GeneralUtils;

public class EWSSubscription implements INotificationEventDelegate,
		ISubscriptionErrorDelegate {

	public static void main(String[] args) throws Exception {
		EWSSubscription sub = new EWSSubscription(1, true);
		sub.connection.addOnNotificationEvent(sub);
		sub.connection.addOnSubscriptionError(sub);
		sub.connection.addOnDisconnect(sub);
		sub.connection.open();
	}

	private StreamingSubscription subscription;

	public StreamingSubscription getSubscription() {
		return subscription;
	}

	public StreamingSubscriptionConnection getConnection() {
		return connection;
	}

	private StreamingSubscriptionConnection connection;

	private boolean reopen;
	private int lifetime;

	public EWSSubscription(int minutes, boolean reopen) throws Exception {
		this.reopen = reopen;
		lifetime = minutes;
		setUpSubscription();
	}

	private void setUpSubscription() throws Exception {
		ArrayList<FolderId> folders = new ArrayList<>();
		folders.add(Folder.bind(ShortCallingService, WellKnownFolderName.Inbox)
				.getId());
		subscription = SubscriptionService.subscribeToStreamingNotifications(
				folders, EventType.Created);
		ArrayList<StreamingSubscription> subscriptions = new ArrayList<>();
		subscriptions.add(subscription);
		connection = new StreamingSubscriptionConnection(SubscriptionService,
				subscriptions, lifetime);

		System.out.println("Ready");
	}

	@Override
	public void notificationEventDelegate(Object sender,
			NotificationEventArgs args) {
		try {
			System.out.println("\nNotification recieved\n");
			for (NotificationEvent event : args.getEvents()) {
				if (event instanceof ItemEvent) {
					ItemEvent e = (ItemEvent) event;
					Item item = Item.bind(ShortCallingService, e.getItemId());
					// System.out.println(MessageBody
					// .getStringFromMessageBody(item.getBody()));
					System.out.println(GeneralUtils.formatHTML(item.getBody()
							.toString()));
				}
			}
			connection.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Override
	public void subscriptionErrorDelegate(Object sender,
			SubscriptionErrorEventArgs args) {
		System.out.println("Disconnected");
		if (reopen)
			try {
				connection.open();
				System.out.println("Reopened");
			} catch (ServiceLocalException e) {
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			}
	}
}

package responseListener;

import static resources.EWSSetup.ShortCallingService;
import static resources.EWSSetup.SubscriptionService;

import java.awt.Event;
import java.util.ArrayList;

import email.OverageNotifications;
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
import resources.GeneralUtils;
import subscription.EWSSubscription;

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
				subscriptions, 1);

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

						System.out.println(GeneralUtils.formatHTML(message
								.getBody().toString()));

						if (message
								.getBody()
								.toString()
								.contains(
										"<font size=\"0\">Sent by Overage Notifier</font>")) {
							System.out
									.println("It's an overage notifiation message");
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

}

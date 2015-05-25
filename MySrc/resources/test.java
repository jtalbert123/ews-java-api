package resources;

import static resources.EWSSetup.ShortCallingService;

import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import responseListener.BasicResponseDetector;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.enumeration.SortDirection;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import email.EmailMessageCreator;
import email.OverageNotifications;

public class test {

	public static void main(String[] args) throws Exception {
		Collection<EmailAddress> members = ShortCallingService.expandGroup(
				"Overage_Notifier@mechdyne.com").getMembers();
		for (EmailAddress addr : members) {
			System.out.println(addr);
		}

		doStuff();

	}

	private static void doStuff() throws Exception {
		BasicResponseDetector brh = new BasicResponseDetector();

		EmailMessageCreator factory = new OverageNotifications();
		// factory.deleteClassMessages(
		// Folder.bind(ShortCallingService, WellKnownFolderName.Root),
		// DeleteMode.HardDelete);
		// System.out.println("Pre-delete");
		EmailMessage email = sendEmail(factory, 1);
		System.out.println("Email sent");
		// System.out.println(EmailMessageUtils.printMessage(email));

		// factory = new EmailMessageCreator(true);
		// email = sendEmail(factory, 2);

		// System.out.println(EmailMessageUtils.printMessage(email));

		// factory.deleteMessage(
		// Folder.bind(ShortCallingService, WellKnownFolderName.Root),
		// DeleteMode.HardDelete, email);
		;
	}

	public static FindItemsResults<Item> listFirstNMessages(
			ExchangeService service, WellKnownFolderName srcFolder, int n)
			throws Exception {
		Folder folder = Folder.bind(service, srcFolder);
		ItemView view = new ItemView(n);
		view.getOrderBy()
				.add(ItemSchema.DateTimeSent, SortDirection.Descending);
		FindItemsResults<Item> findResults = service.findItems(folder.getId(),
				view);

		return findResults;
	}

	/**
	 * In the message, HTML must be used \n becomes a space (' ').
	 * 
	 * @param service
	 * @throws Exception
	 * @throws ServiceLocalException
	 */
	static EmailMessage sendEmail(EmailMessageCreator factory, int version)
			throws Exception, ServiceLocalException {
		Map<String, Object> properties = new HashMap<>();
		properties.put("user.name", "James");
		properties.put("user.address", new EmailAddress("James Talbert",
				"james.talbert@mechdyne.com"));

		// ShortCallingService
		// .setImpersonatedUserId(new ImpersonatedUserId(
		// ConnectingIdType.SmtpAddress,
		// "_Overage_Notifier@mechdyne.com"));
		EmailMessage message = factory.newEmail(properties);
		// message.setSender(new EmailAddress("James Talbert",
		// "james.talbert@mechdyne.com"));
		// message.getToRecipients().clear();
		// message.getToRecipients()
		// .add(new EmailAddress("jtalbert123@gmail.com"));

		message.sendAndSaveCopy();

		assert message.getId() == null;

		List<EmailMessage> set = factory
				.getMessageInstances(Folder.bind(ShortCallingService,
						WellKnownFolderName.SentItems), message);

		EmailMessage result = null;
		for (EmailMessage m : set) {
			result = m;
			System.out.println(m);
		}
		return result;
	}
}
package resources;

import static resources.EWSSetup.ShortCallingService;

import java.util.List;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.enumeration.DeleteMode;
import microsoft.exchange.webservices.data.enumeration.SortDirection;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import email.EmailMessageCreator;
import email.EmailMessageUtils;

public class test {

	public static void main(String[] args) throws Exception {
		doStuff();

	}

	private static void doStuff() throws Exception {
		EmailMessageCreator factory = new EmailMessageCreator(true);
		factory.deleteClassMessages(
				Folder.bind(ShortCallingService, WellKnownFolderName.Root),
				DeleteMode.HardDelete);
		EmailMessage email = sendEmail(ShortCallingService, factory, 1);

		System.out.println(EmailMessageUtils.printMessage(email));

		factory = new EmailMessageCreator(true);
		email = sendEmail(ShortCallingService, factory, 2);

		System.out.println(EmailMessageUtils.printMessage(email));

		factory.deleteMessage(
				Folder.bind(ShortCallingService, WellKnownFolderName.Root),
				DeleteMode.HardDelete, email);
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
	static EmailMessage sendEmail(ExchangeService service,
			EmailMessageCreator factory, int version) throws Exception,
			ServiceLocalException {
		EmailMessage message = factory.newEmail();
		message.setSender(new EmailAddress("James Talbert",
				"james.talbert@mechdyne.com"));
		message.setSubject("EWS message " + version);
		message.setBody(new MessageBody("version " + version));
		message.getToRecipients()
				.add(new EmailAddress("jtalbert123@gmail.com"));

		// Specify when to send the email by setting the value of the extended
		// property.

		message.sendAndSaveCopy();

		assert message.getId() == null;

		List<EmailMessage> set = factory.getMessageInstances(
				Folder.bind(service, WellKnownFolderName.SentItems), message);

		EmailMessage result = null;
		for (EmailMessage m : set) {
			result = m;
			// System.out.println(m.getId());
		}
		return result;
	}
}
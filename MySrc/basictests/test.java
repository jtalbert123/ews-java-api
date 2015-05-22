package basictests;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Set;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.enumeration.ExchangeVersion;
import microsoft.exchange.webservices.data.enumeration.SortDirection;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.enumeration.DeleteMode;

public class test {

	public static void main(String[] args) throws Exception {
		ExchangeService service = setUp();

		doStuff(service);

		service.close();
	}

	private static ExchangeService setUp() throws URISyntaxException {
		ExchangeService service = new ExchangeService(
				ExchangeVersion.Exchange2010_SP2);
		ExchangeCredentials credentials = new WebCredentials(
				"jtalbert@mechdyne.com", "2Pets4us");
		service.setCredentials(credentials);
		service.setUrl(new URI(
				"https://outlook.office365.com/ews/Exchange.asmx"));
		return service;
	}

	private static void doStuff(ExchangeService service) throws Exception {
		EmailMessageCreator factory = new EmailMessageCreator(service, true);
		EmailMessage email = sendEmail(service, factory);

		System.out.println(EmailMessageUtils.printMessage(email));

		factory.close();
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
	private static EmailMessage sendEmail(ExchangeService service,
			EmailMessageCreator factory) throws Exception,
			ServiceLocalException {
		EmailMessage message = factory.newEmail();

		message.setSender(new EmailAddress("James Talbert",
				"james.talbert@mechdyne.com"));

		message.setSubject("EWS message");

		message.setBody(new MessageBody(
				"Alan, please reply to this message letting me know that you recieved it.<br>\tThanks!"));

		 message.getToRecipients().add(
		 new EmailAddress("James Talbert", "jtalbert123@gmail.com"));
		message.getToRecipients().add(
				new EmailAddress("James.talbert@mechdyne.com"));

		// Specify when to send the email by setting the value of the extended
		// property.

		message.sendAndSaveCopy();

		assert message.getId() == null;

		Set<EmailMessage> set = factory.getMessageInstances(
				Folder.bind(service, WellKnownFolderName.SentItems), message);

		EmailMessage result = null;
		for (EmailMessage m : set) {
			result = m;
			// System.out.println(m.getId());
		}
		return result;
	}

	private static EmailMessage getLastSent(ExchangeService service)
			throws Exception {
		Folder folder = Folder.bind(service, WellKnownFolderName.Inbox);
		ItemView view = new ItemView(1);
		view.getOrderBy()
				.add(ItemSchema.DateTimeSent, SortDirection.Descending);
		FindItemsResults<Item> findResults = service.findItems(folder.getId(),
				view);

		EmailMessage message = EmailMessage.bind(service, findResults
				.getItems().get(0).getId());

		System.out.println(message.getSubject());
		System.out.println(message.getDateTimeSent());
		// System.out.println(message.getId());
		return message;
	}

	private static EmailMessage getLastDeleted(ExchangeService service)
			throws Exception {
		Folder folder = Folder.bind(service, WellKnownFolderName.DeletedItems);
		ItemView view = new ItemView(1);
		view.getOrderBy()
				.add(ItemSchema.DateTimeSent, SortDirection.Descending);
		FindItemsResults<Item> findResults = service.findItems(folder.getId(),
				view);

		EmailMessage message = (EmailMessage) findResults.getItems().get(0);

		System.out.println(message.getSubject());
		return message;
	}
}

package resources;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.List;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.enumeration.DeleteMode;
import microsoft.exchange.webservices.data.enumeration.ExchangeVersion;
import microsoft.exchange.webservices.data.enumeration.SortDirection;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

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
		factory.deleteClassMessages(
				Folder.bind(service, WellKnownFolderName.Root),
				DeleteMode.HardDelete);
		EmailMessage email = sendEmail(service, factory, 1);

		System.out.println(EmailMessageUtils.printMessage(email));

		factory = new EmailMessageCreator(service, true);
		email = sendEmail(service, factory, 2);

		System.out.println(EmailMessageUtils.printMessage(email));

		factory.deleteMessage(Folder.bind(service, WellKnownFolderName.Root),
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

	private static EmailMessage getMostRecent(ExchangeService service,
			Folder folder) throws Exception {
		ItemView view = new ItemView(1);
		view.getOrderBy()
				.add(ItemSchema.DateTimeSent, SortDirection.Descending);
		FindItemsResults<Item> findResults = service.findItems(folder.getId(),
				view);

		EmailMessage message = (EmailMessage) findResults.getItems().get(0);

		return message;
	}
}

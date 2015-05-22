package tests;

import java.net.URI;
import java.util.List;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.enumeration.ExchangeVersion;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

import resources.EmailMessageCreator;

public class EmailMessageCreatorTest {

	private static ExchangeService service;

	@BeforeClass
	public static void setUpBeforeClass() throws Exception {
		service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
		ExchangeCredentials credentials = new WebCredentials(
				"jtalbert@mechdyne.com", "2Pets4us");
		service.setCredentials(credentials);
		service.setUrl(new URI(
				"https://outlook.office365.com/ews/Exchange.asmx"));
	}

	@AfterClass
	public static void tearDownAfterClass() throws Exception {
		service.close();
	}

	private EmailMessageCreator factory1;
	private EmailMessageCreator factory2;

	@Before
	public void setUp() throws Exception {
		factory1 = new EmailMessageCreator(service);
		factory2 = new EmailMessageCreator(service);
	}

	@Test
	public void test() {

	}

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

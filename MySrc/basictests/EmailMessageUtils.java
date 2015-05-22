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

public class EmailMessageUtils {

	public static String printMessage(EmailMessage message) throws Exception {
		message.load();

		StringBuilder string = new StringBuilder();

		string.append("ID: ").append(message.getId()).append('\n');
		string.append("From:\t").append(message.getFrom()).append('\n');
		string.append("To:");
		for (EmailAddress address : message.getToRecipients()) {
			string.append("\t").append(address.getAddress()).append('\n');
		}
		string.append("\nCC:");
		for (EmailAddress address : message.getCcRecipients()) {
			string.append("\t").append(address.getAddress()).append('\n');
		}
		string.append("\nBCC:");
		for (EmailAddress address : message.getBccRecipients()) {
			string.append("\t").append(address.getAddress()).append('\n');
		}
		string.append("\n\n").append(message.getSubject()).append('\n');
		string.append('\n').append(message.getBody()).append('\n');
		return string.toString();
	}
}

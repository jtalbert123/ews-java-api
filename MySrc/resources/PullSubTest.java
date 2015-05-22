package resources;

import java.net.URI;
import java.net.URISyntaxException;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.enumeration.ExchangeVersion;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import subscription.ItemListener;
import subscription.PullSubscription;
import subscription.PullSubscription.Change;

public class PullSubTest implements ItemListener {

	ExchangeService service;
	PullSubscription ps;

	public static void main(String[] args) throws Exception {
		PullSubTest pst = new PullSubTest();
	}

	public PullSubTest() throws Exception {
		service = setUp();
		ps = new PullSubscription(service, Folder.bind(service,
				WellKnownFolderName.Inbox), 1000 * 30);
		ps.addListener(this);
	}

	private ExchangeService setUp() throws URISyntaxException {
		ExchangeService service = new ExchangeService(
				ExchangeVersion.Exchange2010_SP2);
		ExchangeCredentials credentials = new WebCredentials(
				"jtalbert@mechdyne.com", "2Pets4us");
		service.setCredentials(credentials);
		service.setUrl(new URI(
				"https://outlook.office365.com/ews/Exchange.asmx"));
		return service;
	}

	@Override
	public void changeEvent(Change c) throws Exception {
		ItemId id = c.ItemId;
		Item item = Item.bind(service, id);

		System.out.println(item.getClass().getSimpleName());
		System.out.println(item.getBody());
	}
}

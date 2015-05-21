package basictests;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.enumeration.ExchangeVersion;

public class test {

	public static void main(String[] args) throws Exception {
		ExchangeService service = new ExchangeService(
				ExchangeVersion.Exchange2010_SP2);
		System.out.println("Created service");

		ExchangeCredentials credentials = new WebCredentials(
				"jtalbert@mechdyne.com", "2Pets4us");
		System.out.println("Created credentials");

		service.setCredentials(credentials);
		System.out.println("Set credentials");

		service.autodiscoverUrl("jtalbert@mechdyne.com");
		System.out.println("Set autoDiscover");

		service.close();
		System.out.println("Closed service");
	}
}

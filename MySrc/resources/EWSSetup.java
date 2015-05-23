package resources;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.List;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.enumeration.ExchangeVersion;
import microsoft.exchange.webservices.data.notification.StreamingSubscriptionConnection;

public class EWSSetup {

	/**
	 * An {@link ExchangeService} object denoted for use with handling
	 * {@link StreamingSubscriptionConnection}s.
	 */
	public static final ExchangeService SubscriptionService = getService();

	/**
	 * An {@link ExchangeService} object denoted for use with short-duration
	 * calls to EWS (blocking calls). Examples: binding operations, creating new
	 * {@link Item Items}/{@link Folder Folders}.
	 */
	public static final ExchangeService ShortCallingService = getService();

	private static URI URL = getURL();

	private static WebCredentials Credentials = getCredentials();

	private static final List<URLChangeListener> URLChangeListeners = new ArrayList<>();

	private static final List<CredentialsChangeListener> CredentialsChangeListeners = new ArrayList<>();

	/**
	 * Always returns a new {@link ExchangeService} object. The objects returned
	 * by this method are not updated when new credentials are given.
	 * 
	 * @return
	 */
	public static ExchangeService getService() {
		ExchangeService service = new ExchangeService(
				ExchangeVersion.Exchange2010_SP2);
		service.setCredentials(getCredentials());
		service.setUrl(getURL());

		return service;
	}

	private static URI getURL() {
		if (URL != null) {
			return URL;
		}
		try {
			return new URI("https://outlook.office365.com/ews/Exchange.asmx");
		} catch (URISyntaxException e) {
			e.printStackTrace();
			return null;
		}
	}

	public static WebCredentials getCredentials() {
		if (Credentials != null)
			return Credentials;
		return new WebCredentials("jtalbert@mechdyne.com", "2Pets4us");
	}

	/**
	 * Sets the credentials of the {@link EWSSetup#ShortCallingService} and
	 * {@link EWSSetup#SubscriptionService} objects.
	 * 
	 * @param newCredentials
	 */
	public static void setCredentials(WebCredentials newCredentials) {
		fireCredentialsChange(Credentials, newCredentials);
		Credentials = newCredentials;
		ShortCallingService.setCredentials(Credentials);
		SubscriptionService.setCredentials(Credentials);
	}

	public static void setURL(URI newURL) {
		fireURLChange(URL, newURL);
		URL = newURL;
		ShortCallingService.setUrl(URL);
		SubscriptionService.setUrl(URL);
	}

	public static URI getCurrentURL() {
		return URL;
	}

	public static WebCredentials getCurrentCredentials() {
		return Credentials;
	}

	public static void addURLChangeListener(URLChangeListener listener) {
		URLChangeListeners.add(listener);
	}

	public static void addCredentialsChangeListener(
			CredentialsChangeListener listener) {
		CredentialsChangeListeners.add(listener);
	}

	public static void removeURLChangeListener(URLChangeListener listener) {
		URLChangeListeners.remove(listener);
	}

	public static void removeCredentialsChangeListener(
			CredentialsChangeListener listener) {
		CredentialsChangeListeners.remove(listener);
	}

	private static void fireURLChange(URI oldURL, URI newURL) {
		for (URLChangeListener listener : URLChangeListeners) {
			listener.URLChanged(newURL, oldURL);
		}
	}

	private static void fireCredentialsChange(WebCredentials oldCredentials,
			WebCredentials newCredentials) {
		for (CredentialsChangeListener listener : CredentialsChangeListeners) {
			listener.CredentialsChanged(newCredentials, oldCredentials);
		}
	}

	public interface URLChangeListener {
		public void URLChanged(URI newURL, URI oldURL);
	}

	public interface CredentialsChangeListener {
		public void CredentialsChanged(WebCredentials newCredentials,
				WebCredentials oldCredentials);
	}
}
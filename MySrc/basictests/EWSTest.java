package basictests;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.net.Authenticator;
import java.net.PasswordAuthentication;
import java.net.URL;
import java.time.ZonedDateTime;
import java.util.Date;
import java.util.GregorianCalendar;

import javax.jws.WebMethod;
import javax.jws.WebParam;
import javax.jws.WebService;
import javax.net.ssl.HttpsURLConnection;
import javax.xml.datatype.XMLGregorianCalendar;

import com.sun.org.apache.xerces.internal.jaxp.datatype.XMLGregorianCalendarImpl;

@WebService
public class EWSTest {
	public EWSTest() {
	}

	private static final String BASIC_AUTHENTICATION = "basic";

	@WebMethod(operationName = "CreateAppointment")
	public String createCalendarItem(
			@WebParam(name = "ExchangeServer") String server,
			@WebParam(name = "DomainUser") final String user,
			@WebParam(name = "DomainPassword") final String password,
			@WebParam(name = "Description") String description,
			@WebParam(name = "Location") String location,
			@WebParam(name = "StartDateTime") XMLGregorianCalendar startDate,
			@WebParam(name = "EndDateTime") XMLGregorianCalendar endDate) {

		System.setProperty("http.auth.preference", BASIC_AUTHENTICATION);
		System.setProperty("http.auth.digest.validateServer", "false");
		System.setProperty("http.auth.digest.validateProxy", "false");

		Authenticator.setDefault(new Authenticator() {

			@Override
			protected PasswordAuthentication getPasswordAuthentication() {
				System.out.println("Authentication...");
				System.out.print("Requesting Prompt: ");
				System.out.println(this.getRequestingPrompt());
				System.out.print("Requesting Scheme: ");
				System.out.println(this.getRequestingScheme());
				System.out.print("Requestor Type: ");
				System.out.println(this.getRequestorType());

				return new PasswordAuthentication(user, password.toCharArray());
			}

		});

		StringBuffer URL = new StringBuffer();
		URL.append("https://");
		URL.append(server);
		URL.append("/ews/Exchange.asmx");
		String result = null;
		StringBuffer request = new StringBuffer();
		request.append("<CreateItem \n");
		request.append("               SendMeetingInvitations=\"SendToNone\"\n");
		request.append("               xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\"\n");
		request.append("               xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n");
		request.append("    <Items>\n");
		request.append("        <t:CalendarItem>\n");
		request.append("            <t:Subject>");
		request.append(description);
		request.append("</t:Subject>\n");
		if (startDate != null) {
			request.append("            <t:Start>");
			request.append(startDate.toXMLFormat());
			request.append("</t:Start>\n");
		}
		if (endDate != null) {
			request.append("            <t:End>");
			request.append(endDate.toXMLFormat());
			request.append("</t:End>\n");
		}
		if (location != null) {
			request.append("            <t:Location>");
			request.append(location);
			request.append("</t:Location>\n");
		}
		request.append("        </t:CalendarItem>\n");
		request.append("    </Items>\n");
		request.append("</CreateItem>\n");

		try {
			// Now we create the output document.
			result = rawRawSoapRequest(URL.toString(), request);
		} catch (Exception e) {
			e.printStackTrace();
		}

		return result;

	}

	private String rawRawSoapRequest(String url, StringBuffer requestString)
			throws Exception {
		// Build the soap envelope.
		StringBuilder builder = new StringBuilder();
		builder.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n");
		builder.append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n");
		builder.append("               xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"\n");
		builder.append("               xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"\n");
		builder.append("               xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\"\n");
		builder.append("               xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n");
		builder.append("<soap:Body>\n");
		builder.append(requestString);
		builder.append("</soap:Body>\n");
		builder.append("</soap:Envelope>\n");
//		System.out.println(builder.toString());

		String currentScheme = System.getProperty("http.auth.preference");
		if (currentScheme != null) {
			System.out.print("Current Scheme: ");
			System.out.println(currentScheme);
		} else {
			System.out.print("Current Scheme: ");
			System.out.println("not set");
		}

		URL ewsURL = new URL(url);
		HttpsURLConnection ewsConn = (HttpsURLConnection) ewsURL
				.openConnection();
		ewsConn.setRequestMethod("POST");
		ewsConn.setRequestProperty("Content-type", "text/xml;utf-8");
		ewsConn.setDoInput(true);
		ewsConn.setDoOutput(true);

		PrintWriter pout = new PrintWriter(new OutputStreamWriter(
				ewsConn.getOutputStream(), "UTF-8"), true);
		pout.print(builder.toString());
		pout.flush();
		pout.close();

		if (ewsConn.getResponseCode() == HttpsURLConnection.HTTP_OK) {
			BufferedReader bin = new BufferedReader(new InputStreamReader(
					ewsConn.getInputStream()));

			StringBuilder result = new StringBuilder();
			String line;
			while ((line = bin.readLine()) != null)
				result.append(line);

			return result.toString();
		} else {
			StringBuilder result = new StringBuilder();
			result.append("Bad post... Code: ");
			result.append(ewsConn.getResponseCode());
			result.append(" - ");
			result.append(ewsConn.getResponseMessage());
			return result.toString();
		}

	}

	public static void main(String[] args) {
		EWSTest test = new EWSTest();
		String result = test.createCalendarItem(
				"outlook.office365.com",
				"jtalbert@mechdyne.com",
				"2Pets4us",
				"test event",
				"HQ",
				new XMLGregorianCalendarImpl(GregorianCalendar
						.from(ZonedDateTime.now())),
				new XMLGregorianCalendarImpl(GregorianCalendar
						.from(ZonedDateTime.now())));

		System.out.println(result);
	}
}

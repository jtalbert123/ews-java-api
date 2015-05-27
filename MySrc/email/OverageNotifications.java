package email;

import java.util.HashMap;
import java.util.Map;

import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.enumeration.Importance;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

public class OverageNotifications extends EmailMessageCreator {

	public static final Map<String, Object> DefaultProperties = getDefaults();

	private static Map<String, Object> getDefaults() {
		Map<String, Object> defaults = new HashMap<>();
		defaults.put("subject", "Data Plan Overage Notification");
		defaults.put("sender.address", new EmailAddress("Overage_Notifier",
				"Overage_Notifier@mechdyne.com"));
		defaults.put("readReciept", false);
		defaults.put("deliveryReciept", true);
		defaults.put("importance", Importance.High);
		defaults.put("user.name", "");
		defaults.put("user.number", "(000) 000-0000");
		defaults.put("user.address", null);
		defaults.put("user.dataUsed", 0.0);
		defaults.put("user.dataPlan", 0.5);
		defaults.put("user.dataOverage", 0.0);
		defaults.put("user.overageCharge", 0.0);
		defaults.put("user.planCharge", 0.0);
		defaults.put("user.otherPersonalCharges", 0.0);

		return defaults;
	}

	public OverageNotifications() throws Exception {
		this(false);
	}

	public OverageNotifications(boolean autoDelete) throws Exception {
		super(autoDelete);
	}

	@Override
	public EmailMessage newEmail(Map<String, Object> properties)
			throws Exception {
		properties = merge(properties, DefaultProperties);

		EmailMessage message = super.newEmail(properties);

		message.setSubject(properties.get("subject"));

		message.setFrom((EmailAddress) properties.get("sender.address"));

		message.setIsReadReceiptRequested((Boolean) properties
				.get("readReciept"));
		message.setIsDeliveryReceiptRequested((Boolean) properties
				.get("deliveryReciept"));

		message.setImportance((Importance) properties.get("inportance"));

		message.setBody(getBody(properties));

		message.getToRecipients().add(
				(EmailAddress) properties.get("user.address"));

		// message.getBccRecipients().add(
		// new EmailAddress("Molly Underwood",
		// "Molly.Underwood@mechdyne.com"));

		return message;
	}

	/**
	 * Merges the maps, the entries in map1 always override in the case of a
	 * conflict.
	 * 
	 * @param map1
	 * @param map2
	 * @return
	 */
	private static Map<String, Object> merge(Map<String, Object> map1,
			Map<String, Object> map2) {
		Map<String, Object> merged = new HashMap<>();

		if (map2 != null)
			for (String key : map2.keySet())
				merged.put(key, map2.get(key));

		if (map1 != null)
			for (String key : map1.keySet())
				merged.put(key, map1.get(key));

		return merged;
	}

	private MessageBody getBody(Map<String, Object> properties) {
		Object otherPersonal = null, name = null, number = null, dataUsed = null;
		Object dataPlan = null, dataOverage = null, overageCharge = null, planCharge = null;

		name = properties.get("user.name");
		number = properties.get("user.number");
		dataUsed = properties.get("user.dataUsed");
		dataPlan = properties.get("user.dataPlan");
		dataOverage = properties.get("user.dataOverage");
		overageCharge = properties.get("user.overageCharge");
		planCharge = properties.get("user.planCharge");
		otherPersonal = properties.get("user.otherPersonalCharges");

		// @FormatterOff

		String html = String
				.format("Hello %s,<br>"
						+ "<br>"
						+ "This billing period you have used %.2f GB of data on your company phone. "
						+ "You have selected a data plan with %.2f GB of data, resulting in a %.2f GB overage. "
						+ "You will be charged $%.2f for the overage in addition to the $%.2f for the data plan.<br>"
						+ "Your total other charges that may need approval come to $%.2f. Total extra charges is: $%.2f<br>"
						+ "Where would you like the charges deducted from?<br>"
						+ "<br>"
						+ "Options:"
						+ "<ul>"
							+ "<li>Expense claim</li>"
							+ "<li>Payroll</li>"
							+ "<li>Approved</li>"
						+ "</ul>"
						+ "Please respond to this message with the sentence "
						+ "\"Deduct from <i>location</i>.\" (with or without quotes) "
						+ "Where <i>location</i> is the item from the preceding list (case insensitive). "
						+ "If the expense is approved, use approved for location (ie \"Deduct from approved.\").<br>"
						+ "<br>"
						+ "You can also say \"Deduct from data overage approved.\", or \"Deduct from $20 approved.\". "
						+ "The most important part is to have the \"Deduct from \" there (with a space after), periods "
						+ "('.') are not allowed in your response. If a period is found, your response "
						+ "will be the part before the period.<br>"
						+ "<br>"
						+ "If no such sentence is found, your response will be ignored.",
						name, (Double) dataUsed, (Double) dataPlan,
						(Double) dataOverage, (Double) overageCharge,
						(Double) planCharge, (Double) otherPersonal, (Double) otherPersonal + (Double) overageCharge);
		String text = "<html>"
				+ "<head>"
					+ "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">"
				+ "</head>"
				+ "<body style=\"padding-bottom:40px\">"
					+ "<div>"
						+ "<p dir=\"auto\" style=\"margin-top:0;margin-bottom:0;\">"
							+ html
						+ "</p>" 
						+ "<br>"
						+ "<br>"
						+ "<br>"
						+ "<br>"
						+ "<font size=\"1\">"
							+ "Sent by Overage Notifier"
							+ "<br>"
							+ "Sent for phone number: " + number
						+ "</font>"
						+ "</div>"
				+ "</body>" + "</html>";
		// @FormatterOn
		MessageBody body = new MessageBody(text);
		// body.setText(text);
		return body;
	}
}

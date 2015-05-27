package email;

import static resources.EWSSetup.ShortCallingService;

import java.io.Closeable;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.enumeration.DeleteMode;
import microsoft.exchange.webservices.data.enumeration.LogicalOperator;
import microsoft.exchange.webservices.data.enumeration.MapiPropertyType;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.ExtendedProperty;
import microsoft.exchange.webservices.data.property.complex.ExtendedPropertyCollection;
import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;
import microsoft.exchange.webservices.data.search.filter.SearchFilter.SearchFilterCollection;

/**
 * A class to simplify the handling of email messages. This class will add three
 * tags to all messages created:<br>
 * <br>
 * a class tag, the name of the class that created the email<br>
 * a creator tag, an integer unique to the instance that created the email<br>
 * a message tag, an integer that uniquely identifies the message to the
 * instance that created it.<br>
 * <br>
 * The class contains utilities to get all messages matching any level of tag.
 * 
 * @author jtalbert
 *
 */
public class EmailMessageCreator implements AutoCloseable {
	private static ExtendedPropertyDefinition classIDProperty = init();
	private static ExtendedPropertyDefinition instanceIDProperty;
	private static ExtendedPropertyDefinition messageIDProperty;

	private static int ID = (int) (System.currentTimeMillis() % (Integer.MAX_VALUE));

	private static ExtendedPropertyDefinition init() {
		try {
			instanceIDProperty = new ExtendedPropertyDefinition(
					UUID.fromString("75A5486F-9267-49ca-9B4E-3D04CA9EC179"),
					"CreatorID", MapiPropertyType.Integer);
			messageIDProperty = new ExtendedPropertyDefinition(
					UUID.fromString("75A5486F-9267-49ca-9B4E-3D04CA9EC180"),
					"MessageID", MapiPropertyType.Integer);

			classIDProperty = new ExtendedPropertyDefinition(
					UUID.fromString("75A5486F-9267-49ca-9B4E-3D04CA9EC181"),
					"CreatedFrom", MapiPropertyType.String);
		} catch (Exception e) {
		}

		return classIDProperty;
	}

	private final int creatorID;
	private HashMap<Integer, EmailMessage> createdMessages;
	private HashMap<EmailMessage, Integer> createdIDs;
	private int messageID;
	private boolean autoDelete;

	public EmailMessageCreator() throws Exception {
		this(false);
	}

	public EmailMessageCreator(boolean autoDelete) throws Exception {
		if (instanceIDProperty == null) {
			init();
		}
		creatorID = getId();

		createdMessages = new HashMap<>();
		createdIDs = new HashMap<>();
		messageID = 0;
		this.autoDelete = autoDelete;
	}

	public EmailMessage newEmail(Map<String, Object> properties)
			throws Exception {
		EmailMessage message = new EmailMessage(ShortCallingService);

		int messageID = getNewMessageId();
		message.setExtendedProperty(classIDProperty, this.getClass().getName());
		message.setExtendedProperty(instanceIDProperty, creatorID);
		message.setExtendedProperty(messageIDProperty, messageID);

		message.getBccRecipients().add(
				new EmailAddress("James.Talbert@mechdyne.com"));

		createdMessages.put(messageID, message);
		createdIDs.put(message, messageID);

		return message;
	}

	public EmailMessage newEmail() throws Exception {
		return newEmail(null);
	}

	private static int getId() {
		return ID++;
	}

	private int getNewMessageId() {
		return messageID++;
	}

	/**
	 * 
	 * @param originalReturn
	 *            a {@link EmailMessage} object previously returned by this
	 *            instance from a call to {@link EmailMessageCreator#newEmail()}
	 * 
	 * @return
	 */
	public int getMessageID(EmailMessage originalReturn) {
		return createdIDs.get(originalReturn);
	}

	public void setAutoDelete(boolean autoDelete) {
		this.autoDelete = autoDelete;
	}

	@Override
	public void close() throws IOException {
		if (autoDelete) {
			try {
				deleteInstanceMessages(Folder.bind(ShortCallingService,
						WellKnownFolderName.Root), DeleteMode.HardDelete);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	public void deleteClassMessages(Folder folder, DeleteMode deleteMode)
			throws Exception {
		for (EmailMessage message : getClassMessages(folder)) {
			message.delete(deleteMode);
		}
	}

	public void deleteInstanceMessages(Folder folder, DeleteMode deleteMode)
			throws Exception {
		for (EmailMessage message : getInstanceMessages(folder)) {
			message.delete(deleteMode);
		}
	}

	public void deleteMessage(Folder folder, DeleteMode deleteMode,
			EmailMessage originalReturn) throws Exception {
		for (EmailMessage message : getMessageInstances(folder,
				getMessageID(originalReturn))) {
			message.delete(deleteMode);
		}
	}

	public void deleteMessage(Folder folder, DeleteMode deleteMode,
			int messageID) throws Exception {
		for (EmailMessage message : getMessageInstances(folder, messageID)) {
			message.delete(deleteMode);
		}
	}

	/**
	 * Gets a set of all messages on the server that were created by any
	 * instance of this class (as determined by {@link #getClass()}).
	 * 
	 * @param folder
	 *            the root folder to search from, sub-folders included
	 * @return
	 * @throws Exception
	 */
	public List<EmailMessage> getClassMessages(Folder folder) throws Exception {
		List<Item> itemList = getClassMessages(folder, new ArrayList<Item>());
		List<EmailMessage> messageList = new ArrayList<>();
		for (Item i : itemList)
			if (i instanceof EmailMessage)
				messageList.add((EmailMessage) i);
		return messageList;
	}

	/**
	 * Gets a set of all messages that were created by this instance that exist
	 * on the server.
	 * 
	 * @param folder
	 *            the root folder to search from, sub-folders included
	 * @return
	 * @throws Exception
	 */
	public List<EmailMessage> getInstanceMessages(Folder folder)
			throws Exception {
		List<Item> itemList = getInstanceMessages(folder, new ArrayList<Item>());
		List<EmailMessage> messageList = new ArrayList<>();
		for (Item i : itemList)
			if (i instanceof EmailMessage)
				messageList.add((EmailMessage) i);
		return messageList;
	}

	/**
	 * Gets a set of all instances of the given message in the given folder. The
	 * {@link EmailMessage} given must be the same object as originally returned
	 * by some call to {@link EmailMessageCreator#newEmail()}.
	 * 
	 * @param folder
	 *            the folder to look for the message in (sub-folders are
	 *            searched recursively)
	 * @param originalReturn
	 *            a {@link EmailMessage} object previously returned by this
	 *            instance from a call to {@link EmailMessageCreator#newEmail()}
	 * @return a set of all messages found in the given folder and all
	 *         sub-folders
	 * @throws Exception
	 */
	public List<EmailMessage> getMessageInstances(Folder folder,
			EmailMessage originalReturn) throws Exception {
		List<Item> itemList = getMessageInstances(folder,
				getMessageID(originalReturn), new ArrayList<Item>());
		List<EmailMessage> messageList = new ArrayList<>();
		for (Item i : itemList)
			if (i instanceof EmailMessage)
				messageList.add((EmailMessage) i);
		return messageList;
	}

	/**
	 * Gets a set of all instances of the given message in the given folder. The
	 * messageID is a value returned from a call to
	 * {@link EmailMessageCreator#getMessageID(EmailMessage)}.
	 * 
	 * @param folder
	 *            the folder to look for the message in (sub-folders are
	 *            searched recursively)
	 * @param messageID
	 *            an integer that uniquely identifies the message to this
	 *            instance, obtained from
	 *            {@link EmailMessageCreator#getMessageID(EmailMessage)}
	 * @return a set of all messages found in the given folder and all
	 *         sub-folders
	 * @throws Exception
	 */
	public List<EmailMessage> getMessageInstances(Folder folder, int messageID)
			throws Exception {
		List<Item> itemList = getMessageInstances(folder, messageID,
				new ArrayList<Item>());
		List<EmailMessage> messageList = new ArrayList<>();
		for (Item i : itemList)
			if (i instanceof EmailMessage)
				messageList.add((EmailMessage) i);
		return messageList;
	}

	public List<Item> getClassMessages(Folder folder, List<Item> list) {
		SearchFilter classFilter = new SearchFilter.IsEqualTo(classIDProperty,
				this.getClass().getName());

		list = EmailMessageUtils.filteredSearch(folder, classFilter, true);

		return list;
	}

	public List<Item> getInstanceMessages(Folder folder, List<Item> list) {

		SearchFilter classFilter = new SearchFilter.IsEqualTo(classIDProperty,
				this.getClass().getName());
		SearchFilter instanceFilter = new SearchFilter.IsEqualTo(
				instanceIDProperty, creatorID);
		SearchFilterCollection filters = new SearchFilterCollection(
				LogicalOperator.And, classFilter, instanceFilter);

		list = EmailMessageUtils.filteredSearch(folder, filters, true);

		return list;
	}

	public List<Item> getMessageInstances(Folder folder, int messageID,
			List<Item> list) {

		SearchFilter classFilter = new SearchFilter.IsEqualTo(classIDProperty,
				this.getClass().getName());
		SearchFilter instanceFilter = new SearchFilter.IsEqualTo(
				instanceIDProperty, creatorID);
		SearchFilter messageFilter = new SearchFilter.IsEqualTo(
				messageIDProperty, messageID);
		SearchFilterCollection filters = new SearchFilterCollection(
				LogicalOperator.And, classFilter, instanceFilter, messageFilter);

		list = EmailMessageUtils.filteredSearch(folder, filters, true);

		return list;
	}

	public boolean wasCreatedByMe(EmailMessage message) throws Exception {
		message.load(new PropertySet(classIDProperty));
		ExtendedPropertyCollection properties = message.getExtendedProperties();
		for (ExtendedProperty property : properties) {
			if (property.getValue().equals(this.getClass().getName())) {
				return true;
			}
		}
		return false;
	}
}

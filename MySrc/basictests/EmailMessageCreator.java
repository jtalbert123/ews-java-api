package basictests;

import java.io.Closeable;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Set;
import java.util.UUID;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.FolderSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.enumeration.DeleteMode;
import microsoft.exchange.webservices.data.enumeration.LogicalOperator;
import microsoft.exchange.webservices.data.enumeration.MapiPropertyType;
import microsoft.exchange.webservices.data.enumeration.SortDirection;
import microsoft.exchange.webservices.data.enumeration.WellKnownFolderName;
import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
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
public class EmailMessageCreator implements Closeable, AutoCloseable {
	private static ExtendedPropertyDefinition classIDProperty;
	private static ExtendedPropertyDefinition instanceIDProperty;
	private static ExtendedPropertyDefinition messageIDProperty;

	private static int ID = (int) (System.currentTimeMillis() % (Integer.MAX_VALUE));

	static void init() throws Exception {
		instanceIDProperty = new ExtendedPropertyDefinition(
				UUID.fromString("75A5486F-9267-49ca-9B4E-3D04CA9EC179"),
				"CreatorID", MapiPropertyType.Integer);
		messageIDProperty = new ExtendedPropertyDefinition(
				UUID.fromString("75A5486F-9267-49ca-9B4E-3D04CA9EC180"),
				"MessageID", MapiPropertyType.Integer);

		classIDProperty = new ExtendedPropertyDefinition(
				UUID.fromString("75A5486F-9267-49ca-9B4E-3D04CA9EC181"),
				"CreatedFrom", MapiPropertyType.String);
	}

	private final int creatorID;
	private ExchangeService service;
	private HashMap<Integer, EmailMessage> createdMessages;
	private HashMap<EmailMessage, Integer> createdIDs;
	private int messageID;
	private boolean autoDelete;

	public EmailMessageCreator(ExchangeService service) throws Exception {
		this(service, false);
	}

	public EmailMessageCreator(ExchangeService service, boolean autoDelete)
			throws Exception {
		if (instanceIDProperty == null) {
			init();
		}
		creatorID = getId();
		this.service = service;

		createdMessages = new HashMap<>();
		createdIDs = new HashMap<>();
		messageID = 0;
		this.autoDelete = autoDelete;
	}

	EmailMessage newEmail() throws Exception {
		EmailMessage message = new EmailMessage(service);

		int messageID = getNewMessageId();
		message.setExtendedProperty(classIDProperty, this.getClass().getName());
		message.setExtendedProperty(instanceIDProperty, creatorID);
		message.setExtendedProperty(messageIDProperty, messageID);

		createdMessages.put(messageID, message);
		createdIDs.put(message, messageID);

		return message;
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
				deleteInstanceMessages(
						Folder.bind(service, WellKnownFolderName.Root),
						DeleteMode.HardDelete);
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
	public Set<EmailMessage> getClassMessages(Folder folder) throws Exception {
		return getClassMessages(folder, new HashSet<EmailMessage>());
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
	public Set<EmailMessage> getInstanceMessages(Folder folder)
			throws Exception {
		return getInstanceMessages(folder, new HashSet<EmailMessage>());
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
	public Set<EmailMessage> getMessageInstances(Folder folder,
			EmailMessage originalReturn) throws Exception {
		return getMessageInstances(folder, getMessageID(originalReturn),
				new HashSet<EmailMessage>());
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
	public Set<EmailMessage> getMessageInstances(Folder folder, int messageID)
			throws Exception {
		return getMessageInstances(folder, messageID,
				new HashSet<EmailMessage>());
	}

	public Set<EmailMessage> getClassMessages(Folder folder,
			Set<EmailMessage> set) {
		SearchFilter classFilter = new SearchFilter.IsEqualTo(classIDProperty,
				this.getClass().getName());
		SearchFilterCollection filter = new SearchFilterCollection(
				LogicalOperator.And, classFilter);

		filteredSearch(folder, set, filter);

		return set;
	}

	public Set<EmailMessage> getInstanceMessages(Folder folder,
			Set<EmailMessage> set) {

		SearchFilter classFilter = new SearchFilter.IsEqualTo(classIDProperty,
				this.getClass().getName());
		SearchFilter instanceFilter = new SearchFilter.IsEqualTo(
				instanceIDProperty, creatorID);
		SearchFilterCollection filters = new SearchFilterCollection(
				LogicalOperator.And, classFilter, instanceFilter);

		filteredSearch(folder, set, filters);

		return set;
	}

	public Set<EmailMessage> getMessageInstances(Folder folder, int messageID,
			Set<EmailMessage> set) {

		SearchFilter classFilter = new SearchFilter.IsEqualTo(classIDProperty,
				this.getClass().getName());
		SearchFilter instanceFilter = new SearchFilter.IsEqualTo(
				instanceIDProperty, creatorID);
		SearchFilter messageFilter = new SearchFilter.IsEqualTo(
				messageIDProperty, messageID);
		SearchFilterCollection filters = new SearchFilterCollection(
				LogicalOperator.And, classFilter, instanceFilter, messageFilter);

		filteredSearch(folder, set, filters);

		return set;
	}

	private void filteredSearch(Folder folder, Set<EmailMessage> set,
			SearchFilterCollection filter) {
		try {
			FindItemsResults<Item> findResults = null;
			do {
				ItemView view = new ItemView(10);
				view.getOrderBy().add(ItemSchema.DateTimeSent,
						SortDirection.Descending);

				findResults = service.findItems(folder.getId(), filter, view);

				for (Item item : findResults.getItems()) {
					if (item instanceof EmailMessage)
						set.add((EmailMessage) item);
				}
			} while (findResults.isMoreAvailable());
		} catch (Exception e) {
		}

		FolderView view = new FolderView(10);
		SearchFilter searchFilter = new SearchFilter.IsGreaterThan(
				FolderSchema.TotalCount, 0);
		try {
			FindFoldersResults results = service.findFolders(folder.getId(),
					searchFilter, view);
			for (Folder sub : results) {
				getClassMessages(sub, set);
			}

		} catch (Exception e) {
		}
	}
}

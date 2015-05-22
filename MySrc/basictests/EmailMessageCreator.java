package basictests;

import java.io.Closeable;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Set;
import java.util.UUID;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
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

public class EmailMessageCreator implements Closeable, AutoCloseable {
	static ExtendedPropertyDefinition creatorIDProperty;
	static ExtendedPropertyDefinition msgIDProperty;
	private static int ID = (int) (System.currentTimeMillis() % (Integer.MAX_VALUE));

	static void init() throws Exception {
		creatorIDProperty = new ExtendedPropertyDefinition(
				UUID.fromString("75A5486F-9267-49ca-9B4E-3D04CA9EC179"),
				"CreatorID", MapiPropertyType.Integer);
		msgIDProperty = new ExtendedPropertyDefinition(
				UUID.fromString("75A5486F-9267-49ca-9B4E-3D04CA9EC180"),
				"MessageID", MapiPropertyType.Integer);
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
		if (creatorIDProperty == null) {
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
		message.setExtendedProperty(creatorIDProperty, creatorID);

		int messageID = getMessageId();

		message.setExtendedProperty(msgIDProperty, messageID);

		createdMessages.put(messageID, message);
		createdIDs.put(message, messageID);

		return message;
	}

	void deleteThisMessages() throws Exception {
		Folder f = Folder.bind(service, WellKnownFolderName.Root);
		deleteThisMessages(f);
	}

	private void deleteThisMessages(Folder f) throws Exception {
		try {
			FindItemsResults<Item> findResults = null;
			do {
				findResults = searchFolder(service,
						Folder.bind(service, f.getId()), 10);

				for (Item item : findResults.getItems()) {
					item.delete(DeleteMode.HardDelete);
				}
			} while (findResults.isMoreAvailable());
		} catch (Exception e) {
		}

		FolderView view = new FolderView(10);
		SearchFilter searchFilter = new SearchFilter.IsGreaterThan(
				FolderSchema.TotalCount, 0);
		FindFoldersResults results = service.findFolders(f.getId(),
				searchFilter, view);

		for (Folder sub : results) {
			deleteThisMessages(sub);
		}
	}

	private static int getId() {
		return ID++;
	}

	private int getMessageId() {
		return messageID++;
	}

	public void deleteMessage(EmailMessage message)
			throws ServiceLocalException, Exception {
		Set<EmailMessage> messages = findMessageInstances(service,
				Folder.bind(service, WellKnownFolderName.Root),
				(int) createdIDs.get(message), new HashSet<EmailMessage>());

		for (EmailMessage inst : messages) {
			inst.delete(DeleteMode.HardDelete);
		}
	}

	public void deleteMessage(int mID) throws ServiceLocalException, Exception {
		Set<EmailMessage> messages = findMessageInstances(service,
				Folder.bind(service, WellKnownFolderName.Root), mID,
				new HashSet<EmailMessage>());

		for (EmailMessage inst : messages) {
			inst.delete(DeleteMode.HardDelete);
		}
	}

	private FindItemsResults<Item> searchFolder(ExchangeService service,
			Folder folder, int limit) throws ServiceLocalException, Exception {
		ItemView view = new ItemView(limit);
		view.getOrderBy()
				.add(ItemSchema.DateTimeSent, SortDirection.Descending);
		SearchFilter idFilter = new SearchFilter.IsEqualTo(creatorIDProperty,
				creatorID);
		FindItemsResults<Item> findResults = service.findItems(folder.getId(),
				idFilter, view);
		return findResults;
	}

	public Set<EmailMessage> getMessage(EmailMessage message, Folder root)
			throws ServiceLocalException, Exception {
		Set<EmailMessage> set = findMessageInstances(service, root,
				createdIDs.get(message), new HashSet<EmailMessage>());
		return set;
	}

	public Set<EmailMessage> getMessage(int messageID, Folder root)
			throws ServiceLocalException, Exception {
		Set<EmailMessage> set = findMessageInstances(service, root, messageID,
				new HashSet<EmailMessage>());
		return set;
	}

	Set<EmailMessage> findMessageInstances(ExchangeService service,
			Folder folder, int mID, Set<EmailMessage> set)
			throws ServiceLocalException, Exception {

		try {
			FindItemsResults<Item> findResults;
			do {
				ItemView view = new ItemView(20);
				view.getOrderBy().add(ItemSchema.DateTimeSent,
						SortDirection.Descending);
				// view.getPropertySet().add();
				SearchFilter creatorIDFilter = new SearchFilter.IsEqualTo(
						creatorIDProperty, creatorID);
				SearchFilter messageIDFilter = new SearchFilter.IsEqualTo(
						msgIDProperty, mID);
				SearchFilter.SearchFilterCollection filters = new SearchFilter.SearchFilterCollection(
						LogicalOperator.Or, creatorIDFilter, messageIDFilter);
				findResults = service.findItems(folder.getId(), filters, view);

				for (Item item : findResults.getItems()) {
					if (item instanceof EmailMessage)
						item.load();
					set.add((EmailMessage) item);
				}
				view.setOffset(view.getOffset() + 20);
			} while (findResults.isMoreAvailable());
		} catch (Exception e) {
		}

		FolderView view = new FolderView(10);
		SearchFilter searchFilter = new SearchFilter.IsGreaterThan(
				FolderSchema.TotalCount, 0);
		FindFoldersResults results;
		do {
			results = service.findFolders(folder.getId(), searchFilter, view);

			for (Folder sub : results) {
				findMessageInstances(service, sub, mID, set);
			}
			view.setOffset(view.getOffset() + 10);
		} while (results.isMoreAvailable());

		return set;
	}

	public int getmID(EmailMessage message) {
		return createdIDs.get(message);
	}

	public void setAutoDelete(boolean autoDelete) {
		this.autoDelete = autoDelete;
	}

	@Override
	protected void finalize() throws Throwable {
		super.finalize();
		close();
	}

	@Override
	public void close() throws IOException {
		if (autoDelete) {
			try {
				deleteThisMessages();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}

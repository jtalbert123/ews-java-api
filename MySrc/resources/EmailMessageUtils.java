package resources;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.FolderSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.enumeration.BasePropertySet;
import microsoft.exchange.webservices.data.enumeration.SortDirection;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

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

	public static List<Item> filteredSearch(ExchangeService service,
			Folder folder, SearchFilter filter, boolean recursive) {
		List<Item> list = new ArrayList<Item>();
		filteredSearchKernel(service, folder, list, filter, recursive);
		Collections.sort(list, new Comparator<Item>() {

			@Override
			public int compare(Item i1, Item i2) {
				try {
					long t1 = Math.max(i1.getDateTimeCreated().getTime(), Math
							.max(i1.getDateTimeReceived().getTime(), i1
									.getDateTimeSent().getTime()));
					long t2 = Math.max(i2.getDateTimeCreated().getTime(), Math
							.max(i2.getDateTimeReceived().getTime(), i2
									.getDateTimeSent().getTime()));

					if (t1 > t2) {
						return -1;
					} else if (t1 < t2) {
						return 1;
					}
				} catch (ServiceLocalException e) {
					e.printStackTrace();
				}
				return 0;
			}

		});
		return list;
	}

	private static void filteredSearchKernel(ExchangeService service,
			Folder folder, List<Item> set, SearchFilter filter,
			boolean recursive) {

		try {
			FindItemsResults<Item> findResults = null;
			PropertySet toDownload = new PropertySet(BasePropertySet.IdOnly);
			ItemView view = new ItemView(10);
			view.getOrderBy().add(ItemSchema.DateTimeSent,
					SortDirection.Descending);
			view.setPropertySet(toDownload);
			do {
				if (filter != null)
					findResults = service.findItems(folder.getId(), filter,
							view);
				else
					findResults = service.findItems(folder.getId(), view);

				for (Item item : findResults.getItems()) {
					set.add(item);
				}
			} while (findResults.isMoreAvailable());
		} catch (Exception e) {
		}

		if (recursive) {
			FolderView view = new FolderView(10);
			SearchFilter searchFilter = new SearchFilter.IsGreaterThan(
					FolderSchema.TotalCount, 0);
			FindFoldersResults results;
			try {
				do {
					results = service.findFolders(folder.getId(), searchFilter,
							view);
					for (Folder sub : results) {
						filteredSearchKernel(service, sub, set, filter, true);
					}
				} while (results.isMoreAvailable());

			} catch (Exception e) {
			}
		}
	}
}

package subscription;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;
import resources.EmailMessageUtils;

public class PullSubscription implements Runnable {

	private long miliseconds;
	private final ExchangeService service;
	private boolean running;
	private Folder folder;

	private List<ItemId> idList;

	private Collection<ItemListener> listeners;

	Thread t;

	public PullSubscription(ExchangeService service, Folder folder,
			long miliSeconds) {
		this.miliseconds = miliSeconds;
		this.service = service;
		this.folder = folder;
		listeners = new ArrayList<>();

		t = new Thread(this);
		// t.start();

		System.out.println("Checking for messages");

		List<ItemId> newList = loadItemIDs();
		List<Change> changes = compare(idList, newList);

		System.out.println(newList.size() + " messages found.");
		System.out.println(changes.size() + " changes found.");
		for (Change c : changes) {
			fireEvent(c);
		}
	}

	@Override
	public void run() {
		if (idList == null || idList.size() == 0) {
			idList = loadItemIDs();
		}
		running = true;

		long oldTime = System.currentTimeMillis();
		while (running) {
			try {
				System.out.println("Checking for messages");

				List<ItemId> newList = loadItemIDs();
				List<Change> changes = compare(idList, newList);

				System.out.println(newList.size() + " messages found.");
				System.out.println(changes.size() + " changes found.");
				for (Change c : changes) {
					fireEvent(c);
				}

				long waitTime = miliseconds
						- (System.currentTimeMillis() - oldTime);
				oldTime = System.currentTimeMillis();
				if (waitTime > 0)
					Thread.sleep(waitTime);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	}

	public void setRunning(boolean running) {
		if (!this.running && running) {
			t.start();
		}
		this.running = running;
	}

	private List<ItemId> loadItemIDs() {

		List<Item> itemList = EmailMessageUtils.filteredSearch(service, folder,
				null, false);
		System.out.println(itemList.size());

		List<ItemId> idList = new ArrayList<>(itemList.size());

		for (Item i : itemList) {
			try {
				idList.add(i.getId());
			} catch (ServiceLocalException e) {
				e.printStackTrace();
			}
		}
		return idList;
	}

	private List<Change> compare(List<ItemId> oldList, List<ItemId> newList) {
		List<Change> changes = new ArrayList<>();

		Iterator<ItemId> iter = oldList.iterator();
		while (iter.hasNext()) {
			ItemId id = iter.next();
			if (!newList.contains(id)) {
				changes.add(new Change(id, ChangeType.REMOVE, oldList
						.indexOf(id)));
				iter.remove();
			}
		}

		iter = newList.iterator();
		while (iter.hasNext()) {
			ItemId id = iter.next();
			if (!oldList.contains(id)) {
				changes.add(new Change(id, ChangeType.ADD, newList.indexOf(id)));
				iter.remove();
			}
		}
		return changes;
	}

	public class Change {
		public final ItemId ItemId;
		public final ChangeType type;
		public final int index;

		public Change(ItemId affected, ChangeType type, int index) {
			ItemId = affected;
			this.type = type;
			this.index = index;
		}
	}

	private enum ChangeType {
		ADD, REMOVE
	}

	public void addListener(ItemListener i) {
		listeners.add(i);
	}

	public void removeListener(ItemListener i) {
		listeners.remove(i);
	}

	private void fireEvent(Change c) {
		for (ItemListener il : listeners) {
			try {
				il.changeEvent(c);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}

package subscription;

import subscription.PullSubscription.Change;

public interface ItemListener {

	void changeEvent(Change c) throws Exception;

}

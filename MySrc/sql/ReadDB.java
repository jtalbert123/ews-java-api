package sql;

import static resources.EWSSetup.ShortCallingService;

import java.sql.ResultSet;
import java.sql.Statement;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Set;

import resources.GeneralUtils;
import microsoft.exchange.webservices.data.misc.NameResolution;
import microsoft.exchange.webservices.data.misc.NameResolutionCollection;

public class ReadDB {

	public static Set<Map<String, Object>> getMessageProperties()
			throws Exception {
		Set<Map<String, Object>> set = new HashSet<>();

		Statement statement = SQLUtilities.getNewStatement();

		ResultSet results = statement
				.executeQuery("SELECT * FROM internal_data;");
		while (results.next()) {
			try {
				String name = GeneralUtils.nameCase(results.getString("name"));

				Map<String, Object> map = new HashMap<>();

				map.put("user.dataOverage", results.getDouble("data_overage"));

				if (((Double) map.get("user.dataOverage")) > 0) {
					map.put("user.name", name.substring(0, name.indexOf(' ')));
					map.put("user.number", results.getLong("number"));
					map.put("user.dataUsed", results.getDouble("data_used"));
					map.put("user.dataPlan",
							Math.max(results.getDouble("data_allocated"), .5));
					map.put("user.overageCharge",
							results.getDouble("overage_charge"));
					map.put("user.planCharge",
							(results.getDouble("data_allocated") - .5) * 10);
					map.put("user.otherPersonalCharges",
							results.getDouble("other_personal_charges"));

					NameResolutionCollection addresses = ShortCallingService
							.resolveName(name);
					NameResolution resolved = null;
					try {
						resolved = addresses.iterator().next();
					} catch (NoSuchElementException e) {
						try {
							String lastName = name.substring(name.indexOf(' '))
									.trim();
							addresses = ShortCallingService
									.resolveName(lastName);
							resolved = addresses.iterator().next();
						} catch (NoSuchElementException e2) {
							String firstName = name.substring(0,
									name.indexOf(' ')).trim();
							addresses = ShortCallingService
									.resolveName(firstName);
							resolved = addresses.iterator().next();
						}
					}
					map.put("user.address", resolved.getMailbox());

					set.add(map);

					System.out.println(resolved.getMailbox());
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return set;
	}
}

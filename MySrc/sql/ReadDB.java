package sql;

import static resources.EWSSetup.ShortCallingService;

import java.sql.ResultSet;
import java.sql.Statement;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Set;

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
				String name = nameCase(results.getString("name"));
				NameResolutionCollection addresses = ShortCallingService
						.resolveName(name);
				NameResolution resolved = null;
				try {
					System.out.println(name);
					resolved = addresses.iterator().next();
				} catch (NoSuchElementException e) {
					try {
						String lastName = name.substring(name.indexOf(' '))
								.trim();
						System.out.println(lastName);
						addresses = ShortCallingService.resolveName(lastName);
						resolved = addresses.iterator().next();
					} catch (NoSuchElementException e2) {
						String firstName = name.substring(0, name.indexOf(' '))
								.trim();
						System.out.println(firstName);
						addresses = ShortCallingService.resolveName(firstName);
						resolved = addresses.iterator().next();
					}
				}

				Map<String, Object> map = new HashMap<>();
				map.put("user.name", name.substring(0, name.indexOf(' ')));
				map.put("user.address", resolved.getMailbox());
				map.put("user.dataUsed", results.getDouble("data_used"));
				map.put("user.dataPlan", results.getDouble("data_allocated"));
				map.put("user.dataOverage", results.getDouble("data_overage"));
				map.put("user.overageCharge",
						results.getDouble("overage_charge"));
				map.put("user.planCharge",
						results.getDouble("data_allocated") * 10);

				set.add(map);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return set;
	}

	private static String nameCase(String name) {
		String[] parts = name.split(" ");
		name = "";
		for (String part : parts) {
			name += Character.toUpperCase(part.charAt(0))
					+ part.substring(1).toLowerCase();
			name += " ";
		}
		return name.trim();
	}
}

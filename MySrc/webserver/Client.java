package webserver;

import java.io.IOException;

import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;

public class Client {

	// public static void main(String[] args) throws IOException {
	// HttpClient client = HttpClients.createDefault();
	// HttpPost method = new HttpPost("http://127.0.0.1:65530/notify");
	//
	// // System.out.println("executing POST");
	// HttpResponse response = client.execute(method);
	// // System.out.println("executed");
	//
	// String responseString = EntityUtils.toString(response.getEntity());
	// System.out.println(responseString);
	// }
}

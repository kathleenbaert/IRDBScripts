package scripts;

import java.util.Random;

public class random {
	
	public static void main (String [] args) {
		
		Random r = new Random();
		
		for(int i = 0; i < 31; i++) {

			System.out.println(r.nextInt(13200));
		}
		
		
	}

}

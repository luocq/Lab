package cn.edu.pku.sce.ThreadDemo;

public class PC {

	public static void main(String[] args) {
		Q q = new Q();
		new Producer(q);
		new Consumer(q);
		System.out.println("Press Control-C to stop.");

	}

}

class Q {
	int n;
	boolean valueset = false;
	synchronized int get() {
		if (!valueset) {
			try {
				wait();
			} catch (InterruptedException e) {
				// TODO: handle exception
			}
		}
		valueset = false;
		notify();
		System.out.println("Got: " + n);
		return n;
	}

	synchronized void put(int n) {
		if (valueset) {
			try {
				wait();
			} catch (InterruptedException e) {
				// TODO: handle exception
			}
		}
		this.n = n;
		valueset = true;
		notify();
		System.out.println("Put: " + n);
	}
}

class Producer implements Runnable {
	Q q;

	Producer(Q q) {
		this.q = q;
		new Thread(this, "Producer").start();
	}

	public void run() {
		int i = 0;
		while (true) {
			q.put(i++);
		}
	}
}

class Consumer implements Runnable {
	Q q;

	Consumer(Q q) {
		this.q = q;
		new Thread(this, "Consumer").start();
	}

	public void run() {
		while (true) {
			q.get();
		}
	}
}
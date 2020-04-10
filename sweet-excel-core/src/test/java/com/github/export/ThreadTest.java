package com.github.export;

import com.sun.xml.internal.messaging.saaj.util.ByteInputStream;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

/**
 * @Author: Vachel Wang
 * @Date: 2020/4/2 7:28 下午
 * @Description: 线程测试
 */
@Slf4j
public class ThreadTest {
	@Test
	public void test1()throws Exception{
		ThreadJob threadJob = new ThreadJob();
		for (int i = 1; i <= 20; i++) {
			Thread thread = new Thread(threadJob,"thread-"+i);
			thread.start();
		}
		Thread.sleep(3 * 60 * 1000);
	}

	class ThreadJob implements Runnable {
		byte[] bytes = new byte[2048];
		ByteInputStream byteInputStream = new ByteInputStream(bytes,2048);
		ThreadLocal<ByteInputStream> threadLocal = ThreadLocal.withInitial(()->{
			System.out.println(bytes.toString());
			return byteInputStream;
		}) ;
		@Override
		public void run() {
			while (true) {
				ByteInputStream inputStream = threadLocal.get();
				try {
					Thread.sleep(3 * 60 * 1000);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
			}
			/*Scanner scanner = new Scanner(System.in);
			String str = scanner.nextLine();*/
		}
	}

}

package net.keepsoft.server;

import javax.annotation.Resource;

import net.keepsoft.dao.BaseDao;

import org.springframework.stereotype.Service;

@Service
public class BaseServer {
	@Resource
	private BaseDao dao;
	
	public  BaseDao getBaseDao(){
		return dao;
	}

	public void test(String id) {
		System.out.println(id);
	}
}

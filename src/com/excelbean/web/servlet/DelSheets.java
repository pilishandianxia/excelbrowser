package com.excelbean.web.servlet;

import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedList;
import java.util.List;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.excelbean.ExcelBrowser;

public class DelSheets extends HttpServlet {
	private final Log log = LogFactory.getLog(DelSheets.class);

	@Override
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws IOException {
		doImpl(request, response);
	}

	@Override
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws IOException {
		doImpl(request, response);
	}

	@SuppressWarnings("unchecked")
	protected void doImpl(HttpServletRequest request, HttpServletResponse response) throws IOException {
		JSONObject returnCodeJsonObject = new JSONObject();
		returnCodeJsonObject.put("returnCode", "Ok");

		String returnCode = returnCodeJsonObject.toJSONString();

		String jsonString = request.getParameter("sheetIds");

		if (StringUtils.isEmpty(jsonString)) {
			returnCodeJsonObject.put("returnCode", "Empty parameter");
			returnCode = returnCodeJsonObject.toJSONString();

			OutputStream out = response.getOutputStream();
			out.write(returnCode.getBytes("UTF-8"));
			out.flush();

			return;
		}

		JSONObject jsonObject = JSONObject.parseObject(jsonString);
		if (jsonObject == null) {
			returnCodeJsonObject.put("returnCode", "Json format error");
			returnCode = returnCodeJsonObject.toJSONString();

			OutputStream out = response.getOutputStream();
			out.write(returnCode.getBytes("UTF-8"));
			out.flush();

			return;
		}

		List<Integer> sheetIds = new LinkedList<Integer>();

		JSONArray jsonArray = jsonObject.getJSONArray("sheetIds");
		for (int i = 0; i < jsonArray.size(); i++) {
			int sheetId = jsonArray.getIntValue(i);

			sheetIds.add(sheetId);

		}

		new ExcelBrowser().delSheets(sheetIds);

		OutputStream out = response.getOutputStream();
		out.write(returnCode.getBytes("UTF-8"));
		out.flush();
	}
}

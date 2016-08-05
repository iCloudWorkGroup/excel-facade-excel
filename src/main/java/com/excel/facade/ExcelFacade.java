package com.excel.facade;



import java.util.List;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelSheet;

import com.excel.model.OnlineExcel;
import com.excel.model.complete.CompleteExcel;
import com.excel.model.complete.ReturnParam;
import com.excel.model.complete.SpreadSheet;



public interface ExcelFacade {
	/**
	 * 通过像素还原excel
	 * 
	 * @param rowBegin
	 *            开始行像素
	 * @param rowEnd
	 *            开始列像素
	 * @return CompleteExcel对象
	 */

	public SpreadSheet openExcel(SpreadSheet spreadSheet,ExcelSheet excelSheet, int rowBegin, int rowEnd,
			ReturnParam returnParam) ;
	
	/**
	 * 通过别名加载excel
	 * 
	 * @return CompleteExcel对象
	 */

	public SpreadSheet openExcelByAlais(SpreadSheet spreadSheet,
			ExcelSheet excelSheet, String rowBeginAlais, String rowEndAlais,
			ReturnParam returnParam);
	
	/**
	 * 通过Workbook转换为CompleteExcel
	 * 
	 * @param excel
	 *            CompleteExcel对象
	 * @return CompleteExcel对象
	 */

	/**
	 * 通过Workbook转换为CompleteExcel
	 * 
	 * @param excel
	 *            CompleteExcel对象
	 * @return CompleteExcel对象
	 */

	//public CompleteExcel getExcel(Workbook workbook);
	/**
	 * 保存excel
	 * 
	 * @param excel
	 *            OnlineExcel对象
	 */

	public void saveOrUpdateExcel(OnlineExcel excel) throws Exception;
	
	/**
	 * 获得所有的OnlineExcel对象
	 * 
	 * @return OnlineExcel对象集合
	 */

	public List<OnlineExcel> getAllExcel();
	/**
	 * 非冻结定位还原excel
	 * 
	 * @param height
	 *            高度
	 * @param returnParam
	 *            返回参数
	 * @return
	 */

	public SpreadSheet positionExcel(ExcelSheet excelSheet,
			SpreadSheet spreadSheet, String height, ReturnParam returnParam) ;
	/**
	 * 通过id获得OnlineExcel对象
	 * 
	 * @param id
	 * @return OnlineExcel
	 */

	public OnlineExcel getExcel(String id);
	
	/**
	 * 通过id获得JsonObject
	 * 
	 * @param id
	 * @return JsonObject
	 */

	public String getExcelByExcelId(String excelId);
	/**
	 * 创建一个默认的excel
	 * 
	 * @param excelId
	 *            excelId
	 */
	public ExcelBook createNewExcel(String excelId);
	
}

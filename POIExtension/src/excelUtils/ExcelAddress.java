package excelUtils;


public class ExcelAddress {
	String address;
	String row;
	String col;
	int rowIndex;
	int colIndex;
	
	public ExcelAddress(String address) throws ExcelException {
		try {
			int[] a = ExcelUtils.convertToIndices(address);
			this.rowIndex = a[0];
			this.colIndex = a[1];
		} catch (ExcelException e) {
			this.rowIndex = ExcelUtils.getExcelRow(address);
			this.colIndex = ExcelUtils.getExcelCol(address);
		}
		this.row = ExcelUtils.convertIntToRow(rowIndex);
		this.col = ExcelUtils.convertIntToCol(colIndex);
		this.address = address;
	}
	public ExcelAddress(int rowIndex, int colIndex) {
		this.row = ExcelUtils.convertIntToRow(rowIndex);
		this.col = ExcelUtils.convertIntToRow(colIndex);
		this.address = this.row+this.col;
		this.rowIndex = rowIndex;
		this.colIndex = colIndex;
	}
}

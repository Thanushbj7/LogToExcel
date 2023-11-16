

	public class TableInfo {
	    private String filePath;
	    private int[] rowNumbers;

	    public TableInfo(String filePath, int[] rowNumbers) {
	        this.filePath = filePath;
	        this.rowNumbers = rowNumbers;
	    }

	    public String getFilePath() {
	        return filePath;
	    }

	    public int[] getRowNumbers() {
	        return rowNumbers;
	    }
	}


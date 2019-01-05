package top.tangshitai.excel.exception;

public class POIException extends Exception{
	private static final long serialVersionUID = 1L;

	public POIException() {
		super();
	}

	public POIException(String msg, Throwable e, boolean arg2, boolean arg3) {
		super(msg, e, arg2, arg3);
	}

	public POIException(String msg, Throwable e) {
		super(msg, e);
	}

	public POIException(String msg) {
		super(msg);
	}

	public POIException(Throwable e) {
		super(e);
	}
}

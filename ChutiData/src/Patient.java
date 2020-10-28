/**
 * @author Danindu
 *Date: 2020 06 16
 *Desc.: A class for a patient object that keeps a record of each patient's id, span,  and number of completes and 
 */

public class Patient {
	private String id;
	private int completes;
	private int span;
	public Patient() {
		this.id = "";
		this.completes = 0;
		this.span = 0;
	}

	public Patient(String id) {
		super();
		this.id = id;
		this.completes = 0;
	}
	
	public int getSpan() {
		return span;
	}

	public void setSpan(int span) {
		this.span = span;
	}

	public String getId() {
		return id;
	}


	public void setId(String id) {
		this.id = id;
	}


	public int getCompletes() {
		return completes;
	}


	public void setCompletes(int completes) {
		this.completes = completes;
	}
	
	@Override
	public String toString() {
		return "Patient [id=" + id + ", completes=" + completes + ", span=" + span + "]";
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub

	}

}
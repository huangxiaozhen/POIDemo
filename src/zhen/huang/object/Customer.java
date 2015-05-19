package zhen.huang.object;

public class Customer
{
	private String id;
	
	private String name;
	
	private String job;
	
	private int age;
	
	
	

	public Customer()
	{
		super();
	}

	public Customer(String id, String name, String job, int age)
	{
		super();
		this.id = id;
		this.name = name;
		this.job = job;
		this.age = age;
	}

	public String getId()
	{
		return id;
	}

	public void setId(String id)
	{
		this.id = id;
	}

	public String getName()
	{
		return name;
	}

	public void setName(String name)
	{
		this.name = name;
	}

	public String getJob()
	{
		return job;
	}

	public void setJob(String job)
	{
		this.job = job;
	}

	public int getAge()
	{
		return age;
	}

	public void setAge(int age)
	{
		this.age = age;
	}
	
	

}

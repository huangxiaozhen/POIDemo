package zhen.huang.object;

public class Student
{
	private String studentNum ;
	
	private String  name ;
	
	private String sex;
	
	private int grade ;
	
	private double height;

	
	
	
	public Student()
	{
		super();
	}

	public Student(String studentNum, String name, String sex, int grade,
			double height)
	{
		super();
		this.studentNum = studentNum;
		this.name = name;
		this.sex = sex;
		this.grade = grade;
		this.height = height;
	}

	public String getStudentNum()
	{
		return studentNum;
	}

	public void setStudentNum(String studentNum)
	{
		this.studentNum = studentNum;
	}

	public String getName()
	{
		return name;
	}

	public void setName(String name)
	{
		this.name = name;
	}

	public int getGrade()
	{
		return grade;
	}

	public void setGrade(int grade)
	{
		this.grade = grade;
	}

	public String getSex()
	{
		return sex;
	}

	public void setSex(String sex)
	{
		this.sex = sex;
	}

	public double getHeight()
	{
		return height;
	}

	public void setHeight(double height)
	{
		this.height = height;
	}
	
	

}

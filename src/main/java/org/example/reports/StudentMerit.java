package org.example.reports;

public class StudentMerit {
    private int studentId;
    private String name;
    private int subjectId;
    private String subject;
    private int classId;
    private String className;
    private int marks;

    public StudentMerit(int studentId, String name, int subjectId, String subject, int classId, String className, int marks) {
        this.studentId = studentId;
        this.name = name;
        this.subjectId = subjectId;
        this.subject = subject;
        this.classId = classId;
        this.className = className;
        this.marks = marks;
    }

    public int getStudentId() {
        return studentId;
    }

    public String getName() {
        return name;
    }

    public int getSubjectId() {
        return subjectId;
    }

    public String getSubject() {
        return subject;
    }

    public int getClassId() {
        return classId;
    }

    public String getClassName() {
        return className;
    }

    public int getMarks() {
        return marks;
    }

    @Override
    public String toString() {
        return "Student " + studentId + " Name: " + name +
                " Subject: " + subject + " Class Id " + classId +
                " Class: " + className + " Marks: " + marks +
                "Subject Id " + subjectId;
    }
}
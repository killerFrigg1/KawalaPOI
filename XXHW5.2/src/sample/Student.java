package sample;

import java.io.Serializable;

/**
 * Created by pwilkin on 30-Nov-17.
 */
public class Student implements Serializable {

    protected String name;
    protected String surname;
    protected String pesel;
    protected String idx;
    protected Double grade;
    protected String gradeDetailed;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSurname() {
        return surname;
    }

    public void setSurname(String surname) {
        this.surname = surname;
    }

    public String getPesel() {
        return pesel;
    }

    public void setPesel(String pesel) {
        this.pesel = pesel;
    }

    public String getIdx() {
        return idx;
    }

    public void setIdx(String idx) {
        this.idx = idx;
    }

    public Double getGrade() {
        return grade;
    }

    public void setGrade(Double grade) {
        this.grade = grade;
    }

    public String getGradeDetailed() {
        return gradeDetailed;
    }

    public void setGradeDetailed(String gradeDetailed) {
        this.gradeDetailed = gradeDetailed;
    }
}

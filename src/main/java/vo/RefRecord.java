/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package vo;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import res.Strings;

/**
 *
 * @author diego
 */
public class RefRecord {
    
    private String specialty;
    private String gender;
    private int age;
    private String country;
    private String ethnicity;
    private String religion;
    private String disability;
    private String sexOr;
    private String trust;
    private String grade;
    private String school;
    private Date refDate;
    private boolean anxiety;
    private boolean carreer;
    private boolean clinSkills;
    private boolean communication;
    private boolean conduct;
    private boolean cultural;
    private boolean exam;
    private boolean healthMental;
    private boolean healthPhysical;
    private boolean language;
    private boolean professionalism;
    private boolean adhd;
    private boolean asd;
    private boolean dyslexia;
    private boolean dyspraxia;
    private boolean srtt;
    private boolean team;
    private boolean time;
    private boolean capability;
    private boolean otherRefReason;
    private boolean caseOpen;

    public RefRecord(String specialty, String gender, int age, String country, String ethnicity, String religion, String disability, String sexOr, String trust, String grade, String school, Date refDate, boolean anxiety, boolean carreer, boolean clinSkills, boolean communication, boolean conduct, boolean cultural, boolean exam, boolean healthMental, boolean healthPhysical, boolean language, boolean professionalism, boolean adhd, boolean asd, boolean dyslexia, boolean dyspraxia, boolean srtt, boolean team, boolean time, boolean capability, boolean otherRefReason, boolean caseOpen) {
        this.specialty = specialty;
        this.gender = gender;
        this.age = age;
        this.country = country;
        this.ethnicity = ethnicity;
        this.religion = religion;
        this.disability = disability;
        this.sexOr = sexOr;
        this.trust = trust;
        this.grade = grade;
        this.school = school;
        this.refDate = refDate;
        this.anxiety = anxiety;
        this.carreer = carreer;
        this.clinSkills = clinSkills;
        this.communication = communication;
        this.conduct = conduct;
        this.cultural = cultural;
        this.exam = exam;
        this.healthMental = healthMental;
        this.healthPhysical = healthPhysical;
        this.language = language;
        this.professionalism = professionalism;
        this.adhd = adhd;
        this.asd = asd;
        this.dyslexia = dyslexia;
        this.dyspraxia = dyspraxia;
        this.srtt = srtt;
        this.team = team;
        this.time = time;
        this.capability = capability;
        this.otherRefReason = otherRefReason;
        this.caseOpen = caseOpen;
    }




    public RefRecord() {
    }


    
    
    
    public String getCountry() {
        return country;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public String getSpecialty() {
        return specialty;
    }

    public void setSpecialty(String specialty) {
        this.specialty = specialty;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getEthnicity() {
        return ethnicity;
    }

    public void setEthnicity(String ethnicity) {
        this.ethnicity = ethnicity;
    }

    public String getReligion() {
        return religion;
    }

    public void setReligion(String religion) {
        this.religion = religion;
    }

    public String getDisability() {
        return disability;
    }

    public void setDisability(String disability) {
        this.disability = disability;
    }

    public String getSexOr() {
        return sexOr;
    }

    public void setSexOr(String sexOr) {
        this.sexOr = sexOr;
    }

    public String getTrust() {
        return trust;
    }

    public void setTrust(String trust) {
        this.trust = trust;
    }

    public String getGrade() {
        return grade;
    }

    public void setGrade(String grade) {
        this.grade = grade;
    }

    public String getSchool() {
        return school;
    }

    public void setSchool(String school) {
        this.school = school;
    }

    public Date getRefDate() {
        return refDate;
    }

    public void setRefDate(Date refDate) {
        this.refDate = refDate;
    }
    
    public boolean isAnxiety() {
        return anxiety;
    }

    public void setAnxiety(boolean anxiety) {
        this.anxiety = anxiety;
    }
    
    public boolean isCapability() {
        return capability;
    }

    public void setCapability(boolean capability) {
        this.capability = capability;
    }

    public boolean isCarreer() {
        return carreer;
    }

    public void setCarreer(boolean carreer) {
        this.carreer = carreer;
    }

    public boolean isClinSkills() {
        return clinSkills;
    }

    public void setClinSkills(boolean clinSkills) {
        this.clinSkills = clinSkills;
    }

    public boolean isCommunication() {
        return communication;
    }

    public void setCommunication(boolean communication) {
        this.communication = communication;
    }

    public boolean isConduct() {
        return conduct;
    }

    public void setConduct(boolean conduct) {
        this.conduct = conduct;
    }

    public boolean isCultural() {
        return cultural;
    }

    public void setCultural(boolean cultural) {
        this.cultural = cultural;
    }

    public boolean isExam() {
        return exam;
    }

    public void setExam(boolean exam) {
        this.exam = exam;
    }

    public boolean isHealthMental() {
        return healthMental;
    }

    public void setHealthMental(boolean healthMental) {
        this.healthMental = healthMental;
    }

    public boolean isHealthPhysical() {
        return healthPhysical;
    }

    public void setHealthPhysical(boolean healthPhysical) {
        this.healthPhysical = healthPhysical;
    }

    public boolean isLanguage() {
        return language;
    }

    public void setLanguage(boolean language) {
        this.language = language;
    }

    public boolean isProfessionalism() {
        return professionalism;
    }

    public void setProfessionalism(boolean professionalism) {
        this.professionalism = professionalism;
    }

    public boolean isAdhd() {
        return adhd;
    }

    public void setAdhd(boolean adhd) {
        this.adhd = adhd;
    }

    public boolean isAsd() {
        return asd;
    }

    public void setAsd(boolean asd) {
        this.asd = asd;
    }

    public boolean isDyslexia() {
        return dyslexia;
    }

    public void setDyslexia(boolean dyslexia) {
        this.dyslexia = dyslexia;
    }

    public boolean isDyspraxia() {
        return dyspraxia;
    }

    public void setDyspraxia(boolean dyspraxia) {
        this.dyspraxia = dyspraxia;
    }

    public boolean isSrtt() {
        return srtt;
    }

    public void setSrtt(boolean srtt) {
        this.srtt = srtt;
    }

    public boolean isTeam() {
        return team;
    }

    public void setTeam(boolean team) {
        this.team = team;
    }

    public boolean isTime() {
        return time;
    }

    public void setTime(boolean time) {
        this.time = time;
    }

    public boolean isOtherRefReason() {
        return otherRefReason;
    }

    public void setOtherRefReason(boolean otherRefReason) {
        this.otherRefReason = otherRefReason;
    }

    public boolean isCaseOpen() {
        return caseOpen;
    }

    public void setCaseOpen(boolean caseOpen) {
        this.caseOpen = caseOpen;
    }
    
    
    
    
    
    public List<String> getAddRef(){
        
        List<String> list = new ArrayList<>();
        
        if(isExam()){
            if(isAnxiety()){
                list.add(Strings.PSW_COLUMN_ANXIETY);
            }
            if(isCarreer()){
                list.add(Strings.PSW_COLUMN_CARREER);
            }
            if(isClinSkills()){
                list.add(Strings.PSW_COLUMN_CLINICAL_SKILLS);
            }
            if(isCommunication()){
                list.add(Strings.PSW_COLUMN_COMMUNICATION);
            }
            if(isConduct()){
                list.add(Strings.PSW_COLUMN_CONDUCT);
            }
            if(isCultural()){
                list.add(Strings.PSW_COLUMN_CULTURAL);
            }
            if(isHealthMental()){
                list.add(Strings.PSW_COLUMN_MENTAL);
            }
            if(isHealthPhysical()){
                list.add(Strings.PSW_COLUMN_PHYSICAL);
            }
            if(isLanguage()){
                list.add(Strings.PSW_COLUMN_LANGUAGE);
            }
            if(isProfessionalism()){
                list.add(Strings.PSW_COLUMN_PROFFESSIONALISM);
            }
            if(isAdhd()){
                list.add(Strings.PSW_COLUMN_ADHD);
            }
            if(isAsd()){
                list.add(Strings.PSW_COLUMN_ASD);
            }
            if(isDyslexia()){
                list.add(Strings.PSW_COLUMN_DYSLEXIA);
            }
            if(isDyspraxia()){
                list.add(Strings.PSW_COLUMN_DYSPRAXIA);
            }
            if(isSrtt()){
                list.add(Strings.PSW_COLUMN_SRTT);
            }
            if(isTeam()){
                list.add(Strings.PSW_COLUMN_TEAM);
            }
            if(isTime()){
                list.add(Strings.PSW_COLUMN_TIME);
            }        
        }
        
        return list;
    }
    
    
}

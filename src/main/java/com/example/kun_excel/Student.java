package com.example.kun_excel;

public class Student {

    String orderNumber; //编号
    String unitName; //单位名称
    String affiliatedGroup; //所属集团
    String isQualified; //是否合格
    String type; //类型
    String rank; //性质
    String checkop; //问题描述
    String checkcontent; //安全防护要求
    String biaozhun; //治理措施
    String problemDetails; //问题详情
    String contactInformation; //人员/联系方式

    public Student(String orderNumber, String unitName, String affiliatedGroup, String isQualified, String type, String rank, String checkop, String checkcontent, String biaozhun, String problemDetails, String contactInformation) {
        this.orderNumber = orderNumber;
        this.unitName = unitName;
        this.isQualified = isQualified;
        this.affiliatedGroup = affiliatedGroup;
        this.type = type;
        this.rank = rank;
        this.checkop = checkop;
        this.checkcontent = checkcontent;
        this.biaozhun = biaozhun;
        this.problemDetails = problemDetails;
        this.contactInformation = contactInformation;
    }
}
package org.example;

import java.util.stream.Collectors;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

// 学生类
class Student {
    private String id;
    private String name;
    private String college; // 新增学院属性
    private String grade;   // 新增年级属性
    private String phone;   // 新增电话属性

    public Student(String id, String name, String college, String grade, String phone) {
        this.id = id;
        this.name = name;
        this.college = college;
        this.grade = grade;
        this.phone = phone;
    }

    // Getter 和 Setter 方法
    public String getCollege() { return college; }
    public String getGrade() { return grade; }
    public String getPhone() { return phone; }
    public String getId() { return id; }
    public String getName() { return name; }

    public void setCollege(String college) { this.college = college; }
    public void setGrade(String grade) { this.grade = grade; }
    public void setPhone(String phone) { this.phone = phone; }
    public void setId(String id) { this.id = id; }
    public void setName(String name) { this.name = name; }
}

// 宿舍类
class Dormitory {
    private String buildingName;
    private int roomNumber;
    private List<Student> students;
    private double rating; // 新增宿舍评分属性

    public Dormitory(String buildingName, int roomNumber) {
        this.buildingName = buildingName;
        this.roomNumber = roomNumber;
        this.students = new ArrayList<>();
        this.rating = 0.0; // 初始评分为 0
    }

    // 添加学生到宿舍
    public void addStudent(Student student) {
        students.add(student);
    }
    public double getRating() {
        return rating;
    }

    public void setRating(double rating) {
        this.rating = rating;
    }

    // Getter 和 Setter
    public String getBuildingName() { return buildingName; }
    public int getRoomNumber() { return roomNumber; }
    public List<Student> getStudents() { return students; }
}

// 宿舍管理系统
class DormitoryManagementSystem {
    private Map<String, Dormitory> dormitories; // 用于存储宿舍信息
    private Map<String, Student> students; // 用于存储学生信息

    public DormitoryManagementSystem() {
        dormitories = new HashMap<>();
        students = new HashMap<>();
    }

    // 添加学生信息
    public void addStudent(Student student) {
        students.put(student.getId(), student);
    }

    // 分配学生到宿舍
    public void assignStudentToDormitory(String studentId, String buildingName, int roomNumber) {
        Student student = students.get(studentId);
        if (student != null) {
            String key = buildingName + "-" + roomNumber;
            Dormitory dormitory = dormitories.get(key);
            if (dormitory != null) {
                dormitory.addStudent(student);
            } else {
                System.out.println("宿舍不存在！");
            }
        } else {
            System.out.println("学生不存在！");
        }
    }

    // 查询学生信息
    public Student queryStudent(String studentId) {
        return students.get(studentId);
    }

    // 添加宿舍信息
    public void addDormitory(Dormitory dormitory) {
        String key = dormitory.getBuildingName() + "-" + dormitory.getRoomNumber();
        dormitories.put(key, dormitory);
    }

    // 查询宿舍信息
    public Dormitory queryDormitory(String buildingName, int roomNumber) {
        return dormitories.get(buildingName + "-" + roomNumber);
    }

    public String findDormitoryOfStudent(String studentId) {
        for (Dormitory dormitory : dormitories.values()) {
            for (Student student : dormitory.getStudents()) {
                if (student.getId().equals(studentId)) {
                    return dormitory.getBuildingName() + " 楼 " + dormitory.getRoomNumber() + " 号室";
                }
            }
        }
        return "学生未分配宿舍或不存在";
    }
    //宿舍评分
    public void setDormitoryRating(String buildingName, int roomNumber, double rating) {
        String key = buildingName + "-" + roomNumber;
        Dormitory dormitory = dormitories.get(key);
        if (dormitory != null) {
            dormitory.setRating(rating);
            System.out.println("宿舍 " + key + " 的评分已更新为 " + rating);
        } else {
            System.out.println("宿舍不存在！");
        }
    }

    // 获取宿舍评分
    public double getDormitoryRating(String buildingName, int roomNumber) {
        String key = buildingName + "-" + roomNumber;
        Dormitory dormitory = dormitories.get(key);
        if (dormitory != null) {
            return dormitory.getRating();
        } else {
            System.out.println("宿舍不存在！");
            return -1; // 表示宿舍不存在
        }
    }
    public void loadStudentsFromExcel(String filePath) throws IOException {
        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // 跳过标题行
                String id = row.getCell(0).getStringCellValue();
                String name = row.getCell(1).getStringCellValue();
                String college = row.getCell(2).getStringCellValue();
                String grade = row.getCell(3).getStringCellValue();
                String phone = row.getCell(4).getStringCellValue();
                addStudent(new Student(id, name, college, grade, phone));
            }
        }
    }

    public void saveStudentsToExcel(String filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fileOut = new FileOutputStream(filePath)) {
            Sheet sheet = workbook.createSheet("Students");
            Row headerRow = sheet.createRow(0);
            String[] columns = {"ID", "Name", "College", "Grade", "Phone", "Dormitory", "Dormitory Rating"};
            for (int i = 0; i < columns.length; i++) {
                headerRow.createCell(i).setCellValue(columns[i]);
            }

            int rowNum = 1;
            for (Student student : students.values()) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(student.getId());
                row.createCell(1).setCellValue(student.getName());
                row.createCell(2).setCellValue(student.getCollege());
                row.createCell(3).setCellValue(student.getGrade());
                row.createCell(4).setCellValue(student.getPhone());

                // 获取学生的宿舍信息
                String dormInfo = findDormitoryOfStudent(student.getId());
                row.createCell(5).setCellValue(dormInfo);

                // 获取宿舍评分
                double dormRating = getDormitoryRatingForStudent(student.getId());
                row.createCell(6).setCellValue(dormRating);
            }
            workbook.write(fileOut);
        }
    }

    private double getDormitoryRatingForStudent(String studentId) {
        for (Dormitory dormitory : dormitories.values()) {
            for (Student student : dormitory.getStudents()) {
                if (student.getId().equals(studentId)) {
                    return dormitory.getRating();
                }
            }
        }
        return 0.0; // 默认评分或学生未分配宿舍
    }

    public void rankDormitoriesByRating() {
        List<Dormitory> sortedDormitories = dormitories.values().stream()
                .sorted((d1, d2) -> Double.compare(d2.getRating(), d1.getRating())) // 按评分降序排序
                .collect(Collectors.toList());

        System.out.println("宿舍评分排名：");
        int rank = 1;
        for (Dormitory dorm : sortedDormitories) {
            System.out.println(rank + ". " + dorm.getBuildingName() + " 楼 " + dorm.getRoomNumber() + " 号室 - 评分: " + dorm.getRating());
            rank++;
        }
    }

}

public class Main {
    public static void main(String[] args) {
        DormitoryManagementSystem system = new DormitoryManagementSystem();
        Scanner scanner = new Scanner(System.in);

        while (true) {
            System.out.println("请选择操作：");
            System.out.println("1. 添加学生");
            System.out.println("2. 查询学生");
            System.out.println("3. 添加宿舍");
            System.out.println("4. 分配宿舍");
            System.out.println("5. 退出");
            System.out.println("6. 从Excel导入学生数据");
            System.out.println("7. 将学生数据导出到Excel");
            System.out.println("8. 设置宿舍评分");
            System.out.println("9. 查询宿舍评分");
            System.out.println("10. 显示宿舍评分排名");
            System.out.print("请输入选项（1-10）：");


            int choice = scanner.nextInt();
            switch (choice) {
                case 1:
                    // 添加学生
                    System.out.print("输入学生ID：");
                    String studentId = scanner.next();
                    System.out.print("输入学生姓名：");
                    String studentName = scanner.next();
                    System.out.print("输入学生学院：");
                    String college = scanner.next();
                    System.out.print("输入学生年级：");
                    String grade = scanner.next();
                    System.out.print("输入学生电话：");
                    String phone = scanner.next();
                    system.addStudent(new Student(studentId, studentName, college, grade, phone));
                    break;

                case 2:
                    // 查询学生
                    System.out.print("输入学生ID：");
                    String id = scanner.next();
                    Student student = system.queryStudent(id);
                    if (student != null) {
                        String studentInfo = "ID: " + student.getId() + ", 姓名: " + student.getName() +
                                ", 学院: " + student.getCollege() + ", 年级: " + student.getGrade() +
                                ", 电话: " + student.getPhone();
                        // 查询学生宿舍信息
                        String dormInfo = system.findDormitoryOfStudent(id);
                        studentInfo += ", 宿舍信息: " + dormInfo;
                        System.out.println(studentInfo);
                    } else {
                        System.out.println("学生不存在！");
                    }
                    break;

                case 3:
                    // 添加宿舍
                    System.out.print("输入宿舍楼名：");
                    String building = scanner.next();
                    System.out.print("输入房间号：");
                    int room = scanner.nextInt();
                    system.addDormitory(new Dormitory(building, room));
                    break;

                case 4:
                    // 分配宿舍
                    System.out.print("输入学生ID：");
                    String sid = scanner.next();
                    System.out.print("输入宿舍楼名：");
                    String bname = scanner.next();
                    System.out.print("输入房间号：");
                    int rnum = scanner.nextInt();
                    system.assignStudentToDormitory(sid, bname, rnum);
                    break;

                case 5:
                    // 退出程序
                    System.out.println("退出程序。");
                    return;

                case 6:
                    // 从Excel导入学生数据
                    System.out.print("输入Excel文件路径：");
                    String importPath = scanner.next();
                    try {
                        system.loadStudentsFromExcel(importPath);
                        System.out.println("学生数据已从Excel导入");
                    } catch (IOException e) {
                        System.out.println("导入Excel文件时出错: " + e.getMessage());
                    }
                    break;

                case 7:
// 将学生数据导出到Excel
                    System.out.print("输入导出Excel文件的路径：");
                    String exportPath = scanner.next();
                    try {
                        system.saveStudentsToExcel(exportPath);
                        System.out.println("学生数据已导出到Excel");
                    } catch (IOException e) {
                        System.out.println("导出到Excel文件时出错: " + e.getMessage());
                    }
                    break;
                case 8:
                    // 设置宿舍评分
                    System.out.print("输入宿舍楼名：");
                    String buildingName = scanner.next();
                    System.out.print("输入房间号：");
                    int roomNumber = scanner.nextInt();
                    System.out.print("输入宿舍评分：");
                    double rating = scanner.nextDouble();
                    system.setDormitoryRating(buildingName, roomNumber, rating);
                    break;

                case 9:
                    // 查询宿舍评分
                    System.out.print("输入宿舍楼名：");
                    buildingName = scanner.next();
                    System.out.print("输入房间号：");
                    roomNumber = scanner.nextInt();
                    double dormRating = system.getDormitoryRating(buildingName, roomNumber);
                    if (dormRating != -1) {
                        System.out.println("宿舍 " + buildingName + " 房间号 " + roomNumber + " 的评分是: " + dormRating);
                    }
                    break;
                case 10:
                    // 显示宿舍评分排名
                    system.rankDormitoriesByRating();
                    break;
            }
        }
    }
}
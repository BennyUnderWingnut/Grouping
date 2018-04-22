import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;

public class Grouping {

    public static void main(String[] args) throws Exception {

        if (args.length != 2) throw new Exception("输入格式为 java -jar Grouping.jar [输入Excel文件] [输出Excel文件]");
        // 从Excel里读取数据
        Workbook workbook = WorkbookFactory.create(new File(args[0]));
        Sheet sheet = workbook.getSheetAt(0);
        ArrayList<Student> students = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();
        int leaders = 0;
        for (Row row : sheet) {
            int ID = Integer.parseInt(dataFormatter.formatCellValue(row.getCell(0)));
            String name = dataFormatter.formatCellValue(row.getCell(1));
            String gender = dataFormatter.formatCellValue(row.getCell(2));
            int leader = Integer.parseInt(dataFormatter.formatCellValue(row.getCell(3)));
            if (leader == 1) leaders++;
            students.add(new Student(ID, name, gender, leader));
        }
        workbook.close();

        // 每6个人至少有一个组长
        int numOfGroups = (int) Math.ceil((double) (students.size()) / 6);
        if (leaders * 6 < students.size()) throw new Exception("组长人数不能小于总人数除以6");

        // 该随机算法在60人左右效果最好，因此先将同学随机分成两份
        Collections.shuffle(students);
        ArrayList<Student>[] studentTypes = new ArrayList[4];
        studentTypes[0] = new ArrayList<>(); //男组长
        studentTypes[1] = new ArrayList<>(); //男组员
        studentTypes[2] = new ArrayList<>(); //女组长
        studentTypes[3] = new ArrayList<>(); //女组员
        ArrayList<Student>[] subClasses = new ArrayList[2];
        subClasses[0] = students;
        subClasses[1] = new ArrayList<>();
        for (Student student : students) {
            if (student.getGender().equals("男")) {
                if (student.getLeader() == 1)
                    studentTypes[0].add(student);
                else studentTypes[1].add(student);
            } else {
                if (student.getLeader() == 1)
                    studentTypes[2].add(student);
                else studentTypes[3].add(student);
            }
        }
        for (int i = 0; i < 4; i++) {
            for (int j = 0; j < studentTypes[i].size() / 2; j++) {
                subClasses[1].add(studentTypes[i].get(j));
                students.remove(studentTypes[i].get(j));
            }
        }

        // 随机生成小组，直到符合条件
        int count = 0;
        for (ArrayList<Student> subClass : subClasses) {
            boolean hasGirl = true;
            boolean hasLeader = true;
            do {
                Collections.shuffle(subClass);
                for (int i = 0; i < (int) Math.ceil((double) (subClass.size()) / 6); i++) {
                    hasGirl = false;
                    hasLeader = false;
                    for (int j = 0; j < 6 && 6 * i + j < subClass.size(); j++) {
                        if (subClass.get(6 * i + j).getLeader() == 1) hasLeader = true;
                        if (subClass.get(6 * i + j).getGender().equals("女")) hasGirl = true;
                    }
                    if (!hasGirl || !hasLeader) break;
                }
            }
            while (!hasGirl || !hasLeader);
        }

        // 第一个subClass里的组数
        int t1 = (int) Math.ceil((double) (subClasses[0].size()) / 6);

        // 写入分组
        Group[] groups = new Group[numOfGroups];
        for (int i = 0; i < numOfGroups; i++) groups[i] = new Group(i);
        for (int i = 0; i < 2; i++) {
            ArrayList<Student> subClass = subClasses[i];
            for (int j = 0; j < subClass.size(); j++) {
                int group = j / 6 + i * t1;
                groups[group].add(subClass.get(j));
            }
        }

        // 对至多五组的五人组进行分配
        if (groups[t1 - 1].size() + groups[groups.length - 1].size() > 6) {
            for (int i = 0; 0 < 5 - groups[t1 - 1].size(); i++) {
                for (Student student : groups[i]) {
                    if (studentTypes[1].contains(student)) {
                        groups[t1 - 1].add(student);
                        groups[i].remove(student);
                        break;
                    }
                }
            }
            for (int i = 0; 0 < 5 - groups[groups.length - 1].size(); i++) {
                for (Student student : groups[t1 + i]) {
                    if (studentTypes[1].contains(student)) {
                        groups[groups.length - 1].add(student);
                        groups[t1 + i].remove(student);
                        break;
                    }
                }
            }
        } else {
            for (Student student : groups[groups.length - 1]) {
                groups[groups.length - 1].remove(student);
                groups[t1 - 1].add(student);
            }
            for (int i = 0; 0 < 5 - groups[t1 - 1].size(); i++) {
                for (Student student : groups[i]) {
                    if (studentTypes[1].contains(student)) {
                        groups[t1 - 1].add(student);
                        groups[t1 + i].remove(student);
                        break;
                    }
                }
            }
        }

        // 写入Excel
        XSSFWorkbook workbook2 = new XSSFWorkbook();
        XSSFSheet sheet2 = workbook2.createSheet("result");
        Row row = sheet2.createRow(0);
        row.createCell(0).setCellValue("组号");
        row.createCell(1).setCellValue("学号");
        row.createCell(2).setCellValue("姓名");
        row.createCell(3).setCellValue("性别");
        row.createCell(4).setCellValue("是否为队长");
        int rowNum = 1;
        for (Group group : groups) {
            for (Student student : group) {
                row = sheet2.createRow(rowNum++);
                row.createCell(0).setCellValue(group.getID() + 1);
                row.createCell(1).setCellValue(student.getID());
                row.createCell(2).setCellValue(student.getName());
                row.createCell(3).setCellValue(student.getGender());
                row.createCell(4).setCellValue(student.getLeader());
            }
        }

        FileOutputStream outputStream = new FileOutputStream(args[1]);
        workbook2.write(outputStream);
        workbook2.close();

        System.out.println("Done");
    }
}

class Student {
    private int ID;
    private String name;
    private String gender;
    private int leader;

    Student(int ID, String name, String gender, int leader) {
        this.ID = ID;
        this.name = name;
        this.gender = gender;
        this.leader = leader;
    }

    int getID() {
        return ID;
    }

    String getName() {
        return name;
    }

    String getGender() {
        return gender;
    }

    int getLeader() {
        return leader;
    }
}

class Group extends ArrayList<Student> {
    private int ID;

    public Group(int ID) {
        super();
        this.ID = ID;
    }

    public int getID() {
        return ID;
    }
}
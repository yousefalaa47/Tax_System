package org.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void main(String[] args) {
        System.out.println("HI");
        // لو عايز تختبر بدون UI
        String testFile = "data.xlsx";
        int testMonth = 2;
        runTax(testFile, testMonth);
    }

    // 🔹 runTax تاخد path للملف والشهر
    public static void runTax(String inputFile, int runMonth) {
        String outputFile = "output_month_" + runMonth + ".xlsx";

        try {
            Connection conn = DriverManager.getConnection(
                    "jdbc:mysql://localhost:3306/tax_db", "root", "Yousef123");

            Statement stmt = conn.createStatement();

            // 🔹 إنشاء الجداول لو مش موجودة
            stmt.execute("CREATE TABLE IF NOT EXISTS AgentData (" +
                    "AgentID INT, AgentName VARCHAR(255), Month INT, Gross DOUBLE, SI DOUBLE, Policy DOUBLE," +
                    "PRIMARY KEY (AgentID, Month))");

            stmt.execute("CREATE TABLE IF NOT EXISTS TaxTable (" +
                    "AgentID INT, Month INT, YTD_Tax DOUBLE," +
                    "PRIMARY KEY (AgentID, Month))");

            // 🔹 مسح بيانات الشهر الحالي فقط
            PreparedStatement delete1 = conn.prepareStatement(
                    "DELETE FROM AgentData WHERE Month=?");
            delete1.setInt(1, runMonth);
            delete1.executeUpdate();

            PreparedStatement delete2 = conn.prepareStatement(
                    "DELETE FROM TaxTable WHERE Month=?");
            delete2.setInt(1, runMonth);
            delete2.executeUpdate();

            // 🔹 قراءة ملف Excel
            FileInputStream fis = new FileInputStream(inputFile);
            Workbook inWorkbook = new XSSFWorkbook(fis);
            Sheet inSheet = inWorkbook.getSheetAt(0);

            // 🔹 إعداد Output Excel
            Workbook outWorkbook = new XSSFWorkbook();
            Sheet outSheet = outWorkbook.createSheet("Result");

            Row header = outSheet.createRow(0);
            header.createCell(0).setCellValue("Agent ID");
            header.createCell(1).setCellValue("Name");
            header.createCell(2).setCellValue("Month");
            header.createCell(3).setCellValue("Monthly Tax");
            header.createCell(4).setCellValue("YTD Tax");

            int outRowIndex = 1;

            for (int i = 1; i <= inSheet.getLastRowNum(); i++) {
                Row row = inSheet.getRow(i);
                if (row == null) continue;

                int agentId = (int) row.getCell(0).getNumericCellValue();
                String name = row.getCell(1).getStringCellValue();
                double gross = row.getCell(2).getNumericCellValue();
                double si = row.getCell(3).getNumericCellValue();
                double policy = row.getCell(4).getNumericCellValue();

                // 🔹 تخزين بيانات الشهر الحالي
                PreparedStatement insert = conn.prepareStatement(
                        "REPLACE INTO AgentData VALUES (?, ?, ?, ?, ?, ?)");
                insert.setInt(1, agentId);
                insert.setString(2, name);
                insert.setInt(3, runMonth);
                insert.setDouble(4, gross);
                insert.setDouble(5, si);
                insert.setDouble(6, policy);
                insert.executeUpdate();

                // 🔹 حساب الضريبة الشهرية بناءً على التراكم
                double monthlyTax = calculateTax(conn, agentId, runMonth);

                // 🔹 الحصول على YTD السابق
                double previousYTD = 0;
                PreparedStatement prev = conn.prepareStatement(
                        "SELECT YTD_Tax FROM TaxTable WHERE AgentID=? AND Month=?");
                prev.setInt(1, agentId);
                prev.setInt(2, runMonth - 1);

                ResultSet rsPrev = prev.executeQuery();
                if (rsPrev.next()) previousYTD = rsPrev.getDouble(1);

                double currentYTD = previousYTD + monthlyTax;

                // 🔹 تخزين YTD
                PreparedStatement save = conn.prepareStatement(
                        "REPLACE INTO TaxTable VALUES (?, ?, ?)");
                save.setInt(1, agentId);
                save.setInt(2, runMonth);
                save.setDouble(3, currentYTD);
                save.executeUpdate();

                // 🔹 كتابة النتيجة في Excel
                Row outRow = outSheet.createRow(outRowIndex++);
                outRow.createCell(0).setCellValue(agentId);
                outRow.createCell(1).setCellValue(name);
                outRow.createCell(2).setCellValue(runMonth);
                outRow.createCell(3).setCellValue(monthlyTax);
                outRow.createCell(4).setCellValue(currentYTD);
            }

            FileOutputStream fos = new FileOutputStream(outputFile);
            outWorkbook.write(fos);

            fos.close();
            outWorkbook.close();
            inWorkbook.close();
            fis.close();
            conn.close();

            System.out.println("🔥 DONE Month " + runMonth);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 🔹 حساب الضريبة الشهرية على أساس التراكم
    private static double calculateTax(Connection conn, int agentId, int runMonth) throws Exception {
        PreparedStatement sum = conn.prepareStatement(
                "SELECT SUM(Gross), SUM(SI), SUM(Policy) FROM AgentData WHERE AgentID=? AND Month<=?");
        sum.setInt(1, agentId);
        sum.setInt(2, runMonth);

        ResultSet rs = sum.executeQuery();

        double totalGross = 0;
        double totalSI = 0;
        double totalPolicy = 0;

        if (rs.next()) {
            totalGross = rs.getDouble(1);
            totalSI = rs.getDouble(2);
            totalPolicy = rs.getDouble(3);
        }

        if (totalPolicy > 833.33 * runMonth)
            totalPolicy = 833.33 * runMonth;

        double taxExemption = 1666.66 * runMonth;
        double actual = totalGross - totalSI - totalPolicy - taxExemption;
        if (actual <= 0) return 0;

        double annualActual = actual * 12 / runMonth;

        double annualTax = calculateAnnualTax(annualActual);

        double taxTillNow = annualTax * runMonth / 12;

        double previousYTD = 0;
        PreparedStatement prev = conn.prepareStatement(
                "SELECT YTD_Tax FROM TaxTable WHERE AgentID=? AND Month=?");
        prev.setInt(1, agentId);
        prev.setInt(2, runMonth - 1);
        ResultSet rsPrev = prev.executeQuery();
        if (rsPrev.next()) previousYTD = rsPrev.getDouble(1);

        return taxTillNow - previousYTD;
    }

    // 🔹 كل شرايح الضريبة السنوية
    private static double calculateAnnualTax(double actual) {
        double tax = 0;

        // First Range
        if (actual <= 600000) {
            double first = actual - 40000;
            if (first <= 0) return 0;
            else if (first <= 15000) tax = first * 0.1;
            else if (first <= 30000) tax = (first - 15000) * 0.15 + 1500;
            else if (first <= 160000) tax = (first - 30000) * 0.2 + 3750;
            else if (first <= 360000) tax = (first - 160000) * 0.225 + 29750;
            else tax = (first - 360000) * 0.25 + 74750;
        }
        // Second Range
        else if (actual <= 700000) {
            if (actual <= 55000) tax = actual * 0.1;
            else if (actual <= 70000) tax = (actual - 55000) * 0.15 + 5500;
            else if (actual <= 200000) tax = (actual - 70000) * 0.2 + 7750;
            else if (actual <= 400000) tax = (actual - 200000) * 0.225 + 33750;
            else tax = (actual - 400000) * 0.25 + 78750;
        }
        // Third Range
        else if (actual <= 800000) {
            if (actual <= 70000) tax = actual * 0.15;
            else if (actual <= 200000) tax = (actual - 70000) * 0.2 + 10500;
            else if (actual <= 400000) tax = (actual - 200000) * 0.225 + 36500;
            else tax = (actual - 400000) * 0.25 + 81500;
        }
        // Fourth Range
        else if (actual <= 900000) {
            if (actual <= 200000) tax = actual * 0.2;
            else if (actual <= 400000) tax = (actual - 200000) * 0.225 + 40000;
            else tax = (actual - 400000) * 0.25 + 85000;
        }
        // Fifth Range
        else if (actual <= 1200000) {
            if (actual <= 400000) tax = actual * 0.225;
            else tax = (actual - 400000) * 0.25 + 90000;
        }
        // Sixth Range
        else {
            if (actual <= 1200000) tax = actual * 0.25;
            else tax = (actual - 1200000) * 0.275 + 300000;
        }

        return tax;
    }
}
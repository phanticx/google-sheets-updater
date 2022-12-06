package com.phanticx;

import com.google.api.client.util.store.MemoryDataStoreFactory;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;
import java.io.*;
import java.util.ArrayList;

import static java.lang.Integer.parseInt;


public class Main {

    private static final String credentialsPath = "PATH_TO_GOOGLE_CREDENTIALS_FILE";
    private static final String excelFilePath = "PATH_TO_EXCEL_FILE";
    private static final JsonFactory jsonFactory = GsonFactory.getDefaultInstance();
    private static final String spreadsheetId = "YOUR_SPREADSHEET_ID";
    private static final String range = "YOUR_SPREADSHEET_SHEET_NAME";


    public static void main(String[] args) {
        try {
            appendSheetsFile(readExcelFile(excelFilePath), readSheetsFile());
        } catch (IOException | GeneralSecurityException e) {
            e.printStackTrace();
        }
    }

    public static List<List<Object>> readExcelFile(String pathName) throws IOException {
        FileInputStream fis = new FileInputStream(new File(pathName));
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        ArrayList<Meeting> meetings = new ArrayList<>();

        for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null && row.getCell(0).getStringCellValue().contains("Meeting #")) {
                ArrayList<User> users = new ArrayList<>();
                String date = sheet.getRow(i + 1).getCell(0).getStringCellValue().substring(17);
                String meetingType = sheet.getRow(i + 2).getCell(0).getStringCellValue().substring(17);
                int userAmt = parseInt(sheet.getRow(i + 3).getCell(0).getStringCellValue().substring(46));
                for (int j = i + 5; j < i + 5 + userAmt; j++) {
                    users.add(new User(sheet.getRow(j).getCell(0).getStringCellValue(), sheet.getRow(j).getCell(1).getStringCellValue().substring(13)));
                }
                meetings.add(new Meeting(users, date, meetingType));
            }
        }

        List<List<Object>> values = new ArrayList<>();
        List<Object> indexValues = new ArrayList<>();

        indexValues.add("Name");
        for (int i = 0; i < meetings.size(); i++) {
            indexValues.add(meetings.get(i).getDate().substring(0,10) + ", " + meetings.get(i).getMeetingType());
        }
        values.add(indexValues);

        List<Object> userList = new ArrayList<>();
        for (Meeting meeting : meetings) {
            for (int i = 0; i < meeting.getUsers().size(); i++) {
                if (!userList.contains(meeting.getUsers().get(i).getName())) {
                    userList.add(meeting.getUsers().get(i).getName());
                }
            }
        }

        for (int i = 0; i < userList.size(); i++) { // for each user i in userList
            List<Object> data = new ArrayList<>();
            data.add(userList.get(i));
            for (int j = 0; j < meetings.size(); j++) { // for each meeting j in meetings
                for (int k = 0; k < meetings.get(j).getUsers().size(); k++) { // for each user k in meeting j
                    if (meetings.get(j).getUsers().get(k).getName().equals(userList.get(i))) {// if user k (in meeting j) is equal to user i (in userList)
                        data.add(meetings.get(j).getDate().substring(0,10));
                    }
                }
                if (data.size() == j + 1) {
                    data.add("Did not attend meeting on " + meetings.get(j).getDate().substring(0,10));
                }
            }
            values.add(data);
        }
        return values;
    }

    public static List<List<Object>> readSheetsFile() throws IOException, GeneralSecurityException {

        ValueRange response = buildSheet().spreadsheets().values()
                         .get(spreadsheetId, range)
                         .execute();
        return response.getValues();
    }

    public static void appendSheetsFile(List<List<Object>> newExcelData, List<List<Object>> oldSheetsData) throws IOException, GeneralSecurityException {
            List<List<Object>> updatedRowsData = oldSheetsData;
            if (newExcelData.size() != oldSheetsData.size()) {  // add new user rows
                List<List<Object>> appendValues = new ArrayList<>();

                for (int i = oldSheetsData.size(); i < newExcelData.size(); i++) { // for new rows i in newExcelData
                    List<Object> newVal = new ArrayList<>();
                    for (int j = 0; j < oldSheetsData.get(0).size(); j++) { // for new cell j in new rows i
                        newVal.add(newExcelData.get(i).get(j));
                    }
                    appendValues.add(newVal);
                    updatedRowsData.add(newVal);
                }


                ValueRange body = new ValueRange()
                        .setValues(appendValues);
                AppendValuesResponse result = buildSheet().spreadsheets().values().append(spreadsheetId, range, body)
                        .setValueInputOption("USER_ENTERED")
                        .setInsertDataOption("INSERT_ROWS")
                        .setIncludeValuesInResponse(true)
                        .execute();
            }

            if (newExcelData.get(0).size() != oldSheetsData.get(0).size()) { // add new columns (meetings)
                List<List<Object>> appendValues = new ArrayList<>();
                System.out.println(updatedRowsData);
                for (int i = 0; i < updatedRowsData.size(); i++) { // for row i in updatedRowsData
                    List<Object> newVal = new ArrayList<>();
                    for (int j = 0; j < newExcelData.size(); j++) { // for row j in newExcelData
                        if (updatedRowsData.get(i).get(0).equals(newExcelData.get(j).get(0))) { // check if same name
                            for (int k = updatedRowsData.get(i).size(); k < newExcelData.get(j).size(); k++) {
                                newVal.add(newExcelData.get(j).get(k));
                            }
                        }
                    }
                    appendValues.add(newVal);
                }

                int columnsAdded = appendValues.get(0).size();
                int oldColumns = oldSheetsData.get(0).size();
                String newRange;

                StringBuilder sb = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();

                if (columnsAdded - oldColumns == 1) { // 1 column added
                    int columnIndex = oldColumns + 1;
                    while (columnIndex-- > 0) {
                        sb.append((char)('A' + (columnIndex % 26)));
                        columnIndex /= 26;
                    }
                    newRange = range + "!" + sb.reverse();
                } else { // more than 1 column added
                    int beginIndex = oldColumns + 1;
                    while (beginIndex-- > 0) {
                        sb.append((char)('A' + (beginIndex % 26)));
                        beginIndex /= 26;
                    }
                    int endIndex = oldColumns + 1;
                    while (endIndex-- > 0) {
                        sb2.append((char)('A' + (endIndex % 26)));
                        endIndex /= 26;
                    }
                    newRange = range + "!" + sb.reverse() + ":" + sb2.reverse();
                }

                ValueRange body = new ValueRange()
                        .setValues(appendValues);
                AppendValuesResponse result = buildSheet().spreadsheets().values().append(spreadsheetId, newRange, body)
                        .setValueInputOption("USER_ENTERED")
                        .setInsertDataOption("OVERWRITE")
                        .setIncludeValuesInResponse(true)
                        .execute();
            }
    }

    public static Sheets buildSheet() throws IOException, GeneralSecurityException {
        InputStream in = new FileInputStream(credentialsPath);
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(jsonFactory, new InputStreamReader(in));

        List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(GoogleNetHttpTransport.newTrustedTransport(), jsonFactory, clientSecrets, scopes)
                                            .setDataStoreFactory(new MemoryDataStoreFactory())
                                            .setAccessType("offline")
                                            .build();
        Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");

        Sheets sheets = new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), jsonFactory, credential)
                .setApplicationName("myapp")
                .build();

        return sheets;
    }

}

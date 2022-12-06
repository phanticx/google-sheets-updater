package com.phanticx;

import java.time.LocalDateTime;
import java.util.ArrayList;

public class Meeting {
    private ArrayList<User> users = new ArrayList<>();
    private String date;
    private String meetingType;
    private LocalDateTime dateTime;

    public Meeting(ArrayList<User> users, String date, String meetingType) {
        this.users = users;
        this.date = date;
        this.meetingType = meetingType;
    }

    public ArrayList<User> getUsers() {
        return users;
    }

    public void setUsers(ArrayList<User> users) {
        this.users = users;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getMeetingType() {
        return meetingType;
    }

    public void setMeetingType(String meetingType) {
        this.meetingType = meetingType;
    }

    public LocalDateTime getDateTime() {
        return dateTime;
    }

    public void setDateTime(LocalDateTime dateTime) {
        this.dateTime = dateTime;
    }
}

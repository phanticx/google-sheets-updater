package com.phanticx;

import java.time.LocalDateTime;

public class User {
    private String name;
    private String loginTime;
    private LocalDateTime loginDateTime;

    public User(String name, String loginTime) {
        this.name = name;
        this.loginTime = loginTime;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getLoginTime() {
        return loginTime;
    }

    public void setLoginTime(String loginTime) {
        this.loginTime = loginTime;
    }

    public LocalDateTime getLoginDateTime() {
        return loginDateTime;
    }

    public void setLoginDateTime(LocalDateTime loginDateTime) {
        this.loginDateTime = loginDateTime;
    }
}

package com.zhsy.web;

import com.zhsy.entity.Role;
import com.zhsy.entity.User;
import com.zhsy.service.UserService;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

/**
 * Created by Archimedes on 2016/5/22.
 */
@Controller
@RequestMapping("/manage")
public class UserController {

    private Logger logger = LoggerFactory.getLogger(this.getClass());

    @Resource
    private UserService userService;

    @RequestMapping(value = "/user")
    public String user(Model model) {
        List<User> users = userService.findAll();
        model.addAttribute("list", users);
        return "user";
    }

    @RequestMapping(value = "/list")
    public String list(Model model) {
        Map<Integer, Object> roleMap = new HashMap<>();
        List<User> users = userService.findAll();
        logger.info("user list = {}", users);
        users.forEach(user -> {
            int userid = user.getId();
            List<Role> roles = userService.findRoles(userid);
            roleMap.put(userid, roles);
        });
        model.addAttribute("users", users);
        model.addAttribute("roleMap", roleMap);
        return "list";
    }
}

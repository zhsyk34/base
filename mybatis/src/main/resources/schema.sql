CREATE TABLE user(
  id INT NOT NULL AUTO_INCREMENT,
  name VARCHAR(20) NOT NULL,
  age INT NOT NULL,
  PRIMARY KEY (id)
)ENGINE = InnoDB AUTO_INCREMENT = 1 DEFAULT CHARSET=UTF8;

CREATE TABLE role(
  id INT NOT NULL AUTO_INCREMENT,
  name VARCHAR(20) NOT NULL,
  PRIMARY KEY (id)
)ENGINE = InnoDB AUTO_INCREMENT = 1 DEFAULT CHARSET=UTF8;

CREATE TABLE userrole(
  userid INT NOT NULL,
  roleid INT NOT NULL
  /*,  CONSTRAINT id PRIMARY KEY (userid, roleid)*/
)ENGINE = InnoDB DEFAULT CHARSET=UTF8;

ALTER TABLE userrole ADD CONSTRAINT id PRIMARY KEY(userid, roleid);
ALTER TABLE userrole ADD CONSTRAINT fk_uid FOREIGN KEY(userid) REFERENCES user(id);
ALTER TABLE userrole ADD CONSTRAINT fk_rid FOREIGN KEY(roleid) REFERENCES role(id);

INSERT INTO user(name, age) VALUES
('zsy',32),
('cjj',32),
('crg',33);

INSERT INTO role(name) VALUES
('高级玩家'),
('超级玩家'),
('骨灰级玩家');

INSERT INTO userrole(userid,roleid) VALUES
(1,1),
(1,3),
(2,1),
(2,2),
(3,1),
(3,2),
(3,3);


CREATE TABLE `demo_user` (
  `id` varchar(32) COLLATE utf8_unicode_ci NOT NULL COMMENT '用户id',
  `name` varchar(16) COLLATE utf8_unicode_ci NOT NULL COMMENT '用户名',
  `gender` bit(1) DEFAULT b'0' COMMENT '性别',
  `age` int(3) unsigned DEFAULT NULL COMMENT '用户年龄',
  `gmt_create` datetime DEFAULT NULL COMMENT '记录创建时间',
  `gmt_modified` datetime DEFAULT NULL COMMENT '记录最近修改时间',
  `deleted` bit(1) DEFAULT b'0' COMMENT '是否删除',
  `phone` varchar(11) COLLATE utf8_unicode_ci DEFAULT NULL COMMENT '电话号码',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_name` (`name`),
  KEY `index_age` (`age`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

insert  into `demo_user`(`id`,`name`,`gender`,`age`,`gmt_create`,`gmt_modified`,`deleted`,`phone`) values ('da5977ce577611ea998680fa5b70e40c','zzs001','\0',18,'2019-10-01 11:13:42','2019-11-01 11:13:50','\0','188******41'),('da597bb6577611ea998680fa5b70e40c','zzs002','\0',18,'2019-11-23 00:00:00','2019-11-23 00:00:00','\0','188******42'),('da597c86577611ea998680fa5b70e40c','zzs003','\0',25,'2019-11-01 11:14:36','2019-11-03 00:00:00','\0','188******43'),('da597ce3577611ea998680fa5b70e40c','zzf001','',26,'2019-11-04 11:14:51','2019-11-03 00:00:00','\0','188******44'),('da597d31577611ea998680fa5b70e40c','zzf002','',17,'2019-11-03 00:00:00','2019-11-03 00:00:00','\0','188******45');


/*
 Navicat Premium Data Transfer

 Source Server         : WorkshopTasks
 Source Server Type    : SQLite
 Source Server Version : 3017000
 Source Schema         : main

 Target Server Type    : SQLite
 Target Server Version : 3017000
 File Encoding         : 65001

 Date: 29/12/2025 14:56:28
*/

PRAGMA foreign_keys = false;

-- ----------------------------
-- Table structure for quality
-- ----------------------------
DROP TABLE IF EXISTS "quality";
CREATE TABLE "quality" (
  "task_id" TEXT NOT NULL,
  "Idex" TEXT DEFAULT NULL,
  "Doc" TEXT DEFAULT NULL,
  "Wiring" TEXT DEFAULT NULL,
  "FunCheck" TEXT DEFAULT NULL,
  "BusSleep" TEXT DEFAULT NULL,
  "TestDrive" TEXT DEFAULT NULL,
  "Cleaning" TEXT DEFAULT NULL,
  "Storing" TEXT DEFAULT NULL,
  "IOcheck" TEXT DEFAULT NULL,
  "HVIBN" TEXT DEFAULT NULL,
  "HVFun" TEXT DEFAULT NULL,
  "SOC" TEXT DEFAULT NULL,
  "Comment" TEXT,
  "Operator" TEXT,
  "Date" TEXT,
  PRIMARY KEY ("task_id")
);

-- ----------------------------
-- Table structure for staff
-- ----------------------------
DROP TABLE IF EXISTS "staff";
CREATE TABLE "staff" (
  "user_account" TEXT,
  "staff_email" TEXT NOT NULL,
  "first_name" TEXT,
  "last_name" TEXT,
  "staff_code" TEXT,
  "interal_external" TEXT,
  "staff_status" TEXT,
  "company" TEXT,
  PRIMARY KEY ("user_account")
);

-- ----------------------------
-- Table structure for staff_archived
-- ----------------------------
DROP TABLE IF EXISTS "staff_archived";
CREATE TABLE "staff_archived" (
  "user_account" TEXT,
  "staff_email" TEXT NOT NULL,
  "first_name" TEXT,
  "last_name" TEXT,
  "staff_code" TEXT,
  "interal_external" TEXT,
  "staff_status" TEXT,
  "company" TEXT,
  PRIMARY KEY ("user_account")
);

-- ----------------------------
-- Table structure for staff_device
-- ----------------------------
DROP TABLE IF EXISTS "staff_device";
CREATE TABLE "staff_device" (
  "user_account" TEXT NOT NULL,
  "computer_SN" TEXT NOT NULL,
  "computer_name" TEXT,
  PRIMARY KEY ("user_account", "computer_SN")
);

-- ----------------------------
-- Table structure for tasks
-- ----------------------------
DROP TABLE IF EXISTS "tasks";
CREATE TABLE "tasks" (
  "task_id" TEXT NOT NULL,
  "task_type" TEXT,
  "suggested_colleague" TEXT,
  "is_finished" TEXT,
  "sender_name" TEXT,
  "sender_dept" TEXT,
  "sender_mail" TEXT,
  "sender_phone" TEXT,
  "cc" TEXT,
  "VIN" TEXT,
  "carid" TEXT,
  "car_model" TEXT,
  "car_location" TEXT,
  "key_location" TEXT,
  "task_status" TEXT,
  "assigned_by" TEXT,
  "accepted_by" TEXT,
  "deferred_by" TEXT,
  "finished_by" TEXT,
  "rejected_by" TEXT,
  "reject_reason" TEXT,
  "sent_time" TEXT,
  "request_start_time" TEXT,
  "request_end_time" TEXT,
  "assigned_time" TEXT,
  "accepted_time" TEXT,
  "finished_time" TEXT,
  "purpose" TEXT,
  "project" TEXT,
  "task_description" TEXT,
  "bg_description" TEXT,
  "effort" REAL,
  "task_dept" TEXT,
  "additional_comment" TEXT,
  PRIMARY KEY ("task_id")
);

-- ----------------------------
-- Table structure for tasks_staff
-- ----------------------------
DROP TABLE IF EXISTS "tasks_staff";
CREATE TABLE "tasks_staff" (
  "task_id" TEXT NOT NULL,
  "staff_email" TEXT NOT NULL,
  "user_account" TEXT,
  "first_name" TEXT,
  "last_name" TEXT,
  "staff_time" REAL,
  PRIMARY KEY ("task_id", "staff_email")
);

PRAGMA foreign_keys = true;

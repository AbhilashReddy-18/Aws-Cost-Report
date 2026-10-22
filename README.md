# Aws-Cost-Report
This project is to understand the cost occurring for individual account in AWS

AWS Cost Intelligence Platform
A serverless AWS cost reporting solution that generates monthly AWS cost reports and delivers them via email.

Architecture
EventBridge ↓ Lambda (Python) ↓ AWS Cost Explorer ↓ Report Generation ↓ S3 ↓ SES Email

Technology Stack
AWS Lambda
EventBridge Scheduler
AWS Cost Explorer
Amazon SES
Amazon S3
Terraform
GitHub Actions
Python

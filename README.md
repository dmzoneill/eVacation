# eVacation

<p align="center">
  <img src=".github/logo.png" alt="eVacation logo" width="128" />
</p>

A web application for employee vacation and leave management.

## What it does

- Employees request leave and see their leave summary, including
  compensatory leave.
- Approvers approve or revoke requests and can delegate approval.
- Administrators grant leave, manage leave types and public holidays,
  and view a holiday calendar.
- Includes query pages, a help page and an admin guide.

## Tech

- Classic ASP (VBScript) on IIS
- Microsoft SQL Server, with scripts in `DBScripts`

## Setup

1. Create the database with the scripts in `DBScripts`:
   `createdatabase.sql`, `createobjects.sql` and
   `fkconstraints.sql`.
2. Host the folder in IIS as the `/eVacation` virtual directory, since
   pages include `/eVacation/common/appglobal.asp`.

## License

Licensed under the Apache License 2.0. See [LICENSE](LICENSE).

# SharePointOnline-MigrationTool

User stories :

# Sign in screen

One-time login

Credential Management

# Source selection screen

Load a recursive tree view from path

Enable item to migrate from selection with fluent interface
E.g. check a folder to select all inner content 

Pre-flight check for SharePoint rules limitation
E.g. Special character, path length, extension type

Pre-designed custom check rules
E.g. Regex check on naming / bytes

Show check results overview inside the app
E.g. % file in compliance, amount, top 3 error category

Write results file on pre-flight check
E.g. Error Details, category

Pre-flight results though PowerBI

# Target selection screen

SPO Site Library selection Tree view expanding to list/lib

Show Site properties on expansion

-Library is empty => Onshot migration scenario

Write results file for Onshot migration

Oneshot results file on powerBI

-Library has root item that match source root items => Delta migration scenario 

Check on modified date to figure what to migrate and write pre-flight results

Write result file on delta scenario

Delta scenario results on powerBI

# Migration history tenant wide

Team modern site to store migration results ??????

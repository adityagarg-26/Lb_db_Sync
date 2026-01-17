# Lb_db_Sync
This process standardizes rate list and export files by removing inactive tests, correcting headers, and matching test names via Ram Baan mapping. Rate list codes and Test Pricing IDs are verified, errors fixed, new tests mapped to master data, and a final Excel is generated and processed.

# Steps
1. Rate List Conversion and correction (Profiles Removal, test names split)
2. Export Download and remove inactive tests.
3. Headers of Rate List and export file changed.
   A. Rate List Header of "Test Name" and rest columns as it is.
   B. Export File Header of "Test Pricing Id", "Test Name Id", "Test", "Master_Test_Name", "Test Type", "Status", "Vendor Price", "Vendor Lb Discount", "Vendor App Discount", delete rest of the columns.
4. Rate List and Export files run in Ram Baan.
5. In Ram Baan Mapping file, Copy paste "Matched_Test" and "Mapped_Test_Name" columns on the right side of "Test Name" column for "Matched" and "Fuzzy Match" Sheet.
6. Check the Test Name and Matched_Test columns in fuzzy match if not correct, then mark them.
7. Put correct ones in fuzzy match to matched sheet.
8. Deleted all columns on the right side of "Test" column from the matched sheet.

9. Copy paste "Sr. No." / "Test code" column as the last column.
10. Bring Rate list code column from ram baan mapping to exported file through vlookup.
11. Put Test Pricing Id column on the right side of Rate List Code column in exported file.
12. In Rate List, bring Test Pricing Id column for the Sr. no. / Test code as the last column through vlookup
13. Put the filter for NA values in exported file and put the rate list codes manually from the rate list for the rest of the tests.
14. Check duplicates in rate list code column in exported file and correct them if showing wrong.
15. Check the prices on the basis of rate list codes.
16. Check NA values in rate list for test pricing id column and take decisions whether to remove in packages or deleted sheet or add as new tests.
17. Map the new add tests with the correct master test names and put the rate list code and rest of the columns in the exported sheet.
18. Check the Master test names through Selected_master_file and take decisions if correctly mapped and if not correctly mapped on the basis of "updated 206 names" then correct them.
19. Special check on these 3 tests, "Quadruple Marker", "Procalcitonin", "Immunofixation Electrophoresis Serum" as they will not show directly as false due to case sensitive.
20. Make a new excel with the headers of Final_War_Code and put all the required column values in that excel.
21. Run the final file through Final_War_Code.
22. Put the new column on the right side of Master Test name with the name, "Rate List Name" and bring it from rate list on the basis of test pricing ids and put the new add ones rate list name by copy pasting.
23. Save the file and send it to Nikita with the test pricing ids that needs to be turned off on the outlook.

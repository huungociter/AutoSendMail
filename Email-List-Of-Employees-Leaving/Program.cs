using EmployeeSendMailModules;
using EmployeeSendMailModels;
using EmployeeSendMailProcessor;

//Call the API and save the employees data into the Employee class.
Employee employeesResponse = await new EmployeeModule().GetEmployeesAsync();

//Filter employees leaving today, create excel file, and send an email. 
EmployeeProcessor employeeProcessor = new EmployeeProcessor();
employeeProcessor.ProcessEmployees(employeesResponse);                        
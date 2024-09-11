using EmployeeSendMailModels;
using System.Net.Http.Json;

namespace EmployeeSendMailModules
{
    public class EmployeeModule
    {
        HttpClient httpClient = new HttpClient();

        public async Task<Employee?> GetEmployeesAsync()
        {
            // Load environment variables from the .env file
            string folderPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + '\\';
            DotNetEnv.Env.Load(folderPath);

            string apiUrl = DotNetEnv.Env.GetString("EMPLOYEE_API_URL");
            return await httpClient.GetFromJsonAsync<Employee>(apiUrl);
        }
    }
}

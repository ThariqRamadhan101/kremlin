# KGB (Kenaikan Gaji Berkala) Upload Automation Script

This script automates the process of uploading KGB documents to the EHRM system using Selenium WebDriver with Chrome debugger connection for enhanced debugging and manual intervention capabilities.

## 🚀 Chrome Debugger Setup Guide

### Why Use Chrome Debugger Mode?

- **Real-time visualization**: Watch exactly what the automation is doing
- **Manual intervention**: Fix issues manually and resume automation
- **Enhanced debugging**: Use Chrome DevTools to inspect elements
- **Interactive testing**: Test XPath selectors in real-time
- **Network monitoring**: Monitor requests and console errors

### Step-by-Step Setup

#### 1. Preparation

- Close all existing Chrome browser windows
- Ensure Chrome is installed on your system
- Make sure ChromeDriver is compatible with your Chrome version

#### 2. Start Chrome with Remote Debugging

**Option A: Command Prompt**

```cmd
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\temp\chrome_debug"
```

**Option B: Chrome Shortcut**

1. Right-click Chrome desktop shortcut
2. Properties > Target
3. Add flags after chrome.exe:

```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\temp\chrome_debug"
```

#### 3. Manual Setup

- Chrome will open with a fresh profile
- Navigate to https://ehrm.pu.go.id/
- Login manually if needed (for testing)
- Keep this Chrome window open

#### 4. Run the Script

- Run the Python script: `python upload_kgb.py`
- It will connect to the existing Chrome session
- You can watch the automation in real-time
- Intervene manually if something goes wrong

#### 5. Debugging Benefits

- **Real-time visualization** of automation steps
- **Chrome DevTools** (F12) to inspect elements
- **Manual fixes** and resume automation
- **Interactive XPath testing**
- **Network monitoring** and console error tracking

## 🔧 Troubleshooting

### Common Issues

**Connection Failed**

- Ensure Chrome is running with `--remote-debugging-port=9222`
- Check if port 9222 is being used by another application
- Try a different port number (update in both Chrome launch and script)

**Elements Not Found**

- Use Chrome DevTools (F12) to verify XPath selectors
- Check if page is fully loaded before automation starts
- Verify element selectors are up-to-date with website changes

**ChromeDriver Issues**

- Ensure ChromeDriver version matches Chrome browser version
- Download compatible ChromeDriver from [Chrome for Testing](https://googlechromelabs.github.io/chrome-for-testing/)

**File Upload Issues**

- Verify KGB files are in the correct folder structure
- Check file names match employee names in Excel
- Ensure PDF files are not corrupted

## 📋 Dependencies

Install required packages:

```bash
pip install -r requirements.txt
```

- **selenium**: WebDriver automation
- **pandas**: Excel file processing
- **python-dotenv**: Environment variable management
- **ChromeDriver**: WebDriver for Chrome (compatible with Chrome version)

## ⚙️ Environment Variables

Create a `.env` file in the same directory:

```env
# EHRM Credentials
USERNAME_EHRM=your_username
PASSWORD_EHRM=your_password

# Date Configuration
SK_DATE_XPATH=//button[contains(@class, 'vs__dropdown-option') and text()='15']
SK_DATE_PREV_MONTH_CLICKS=1
EFFECTIVE_DATE_XPATH=//button[contains(@class, 'vs__dropdown-option') and text()='1']
EFFECTIVE_DATE_NEXT_MONTH_CLICKS=1

# Verification
EXPECTED_USERNAME=your_username_for_verification
```

### Environment Variable Descriptions

- **USERNAME_EHRM**: Your EHRM system username
- **PASSWORD_EHRM**: Your EHRM system password
- **SK_DATE_XPATH**: XPath selector for SK date in datepicker
- **SK_DATE_PREV_MONTH_CLICKS**: Number of times to click previous month for SK date
- **EFFECTIVE_DATE_XPATH**: XPath selector for effective date in datepicker
- **EFFECTIVE_DATE_NEXT_MONTH_CLICKS**: Number of times to click next month for effective date
- **EXPECTED_USERNAME**: Username to verify successful submission

## 📁 File Structure

```
upload_kgb/
├── upload_kgb.py              # Main automation script
├── requirements.txt          # Python dependencies
├── .env                      # Environment variables
├── README.md                 # This file
├── kgb/                      # KGB documents folder
│   ├── Direktorat Sistem.../
│   │   └── 2026_Mei_Saflir...pdf
│   ├── Satuan Kerja Pelaksanaan.../
│   │   └── 2026_Mei_Dian...pdf
│   └── ...
└── kgb_processed_data_*.xlsx # Processed Excel files
```

## 🚦 Usage

1. **Setup Chrome Debugger** (see above)
2. **Prepare Environment**:
   - Install dependencies
   - Create `.env` file with credentials
   - Organize KGB files in `kgb/` folder
3. **Run Automation**:
   ```bash
   python upload_kgb.py
   ```
4. **Monitor Progress**:
   - Watch automation in Chrome
   - Check console output for status
   - Intervene manually if needed

## 🔍 Chrome DevTools Tips

During automation, use Chrome DevTools (F12) to:

- **Elements Tab**: Inspect HTML structure and test XPath selectors
- **Console Tab**: Monitor JavaScript errors and network requests
- **Network Tab**: Track API calls and responses
- **Sources Tab**: Debug page scripts if needed

### Testing XPath Selectors

In DevTools Console:

```javascript
// Test XPath selector
$x('//button[@class="btn btn-primary"]');

// Check element properties
$x('//button[@class="btn btn-primary"]')[0].textContent;
```

## 🛠️ Advanced Configuration

### Custom Chrome Options

Edit the `open_website()` function to modify Chrome behavior:

```python
options = Options()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
# Add custom options
options.add_argument("--disable-notifications")
options.add_argument("--disable-popup-blocking")
```

### Alternative Debug Ports

Change port if 9222 is in use:

1. Update Chrome launch command:

   ```cmd
   chrome.exe --remote-debugging-port=9223 --user-data-dir="C:\temp\chrome_debug"
   ```

2. Update script:
   ```python
   options.add_experimental_option("debuggerAddress", "127.0.0.1:9223")
   ```

## 📝 Script Functions

The script includes these main functions:

- `open_website()`: Connect to Chrome debugger and open EHRM
- `click_enter_button()`: Click 'Masuk' button
- `click_sso_login_button()`: Click 'Login SSO' button
- `fill_username()` / `fill_password()`: Fill login credentials
- `click_login_button()`: Submit login form
- `open_employee_service_page()`: Navigate to employee service
- `read_all_employees_from_excel()`: Read employee data from Excel
- `fill_nip()` / `click_search_button()`: Search for employee
- `click_view_employee_button()`: View employee details
- `click_salary_button()`: Navigate to salary section
- `click_add_data_button()`: Add new salary data
- `fill_*()`: Various form filling functions
- `upload_kgb_file()`: Upload KGB document
- `verify_submission_success()`: Verify successful submission

## 🐛 Debug Mode

The script includes several debugging features:

- **Input prompts**: Pause execution for manual intervention
- **Detailed logging**: Step-by-step progress updates
- **Error handling**: Graceful error recovery with user prompts
- **Chrome connection**: Real-time browser control

## 📞 Support

If you encounter issues:

1. Check Chrome debugger setup
2. Verify environment variables
3. Inspect elements with Chrome DevTools
4. Review console output for error messages
5. Test XPath selectors manually

## 📄 License

This script is for internal use only. Please ensure you have proper authorization before automating interactions with the EHRM system.

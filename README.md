# VoltasNonStandardExpenseEntryBot :robot:

[![VERSION](https://img.shields.io/badge/VERSION-v2.0.43-sucess)](https://github.com/charan379/VoltasNonStandardExpenseEntryBot) [![LAST UPDATE](https://img.shields.io/badge/LAST--UPDATED-07--April--2023-sucess)](https://github.com/charan379/VoltasNonStandardExpenseEntryBot) [![AGPL License](https://img.shields.io/badge/LICENSE-GNU%20AGPLv3-informational)](https://www.gnu.org/licenses/agpl-3.0.en.html)

## About

VoltasNonStandardExpenseEntryBot is developed using UiPath RE-Framework to automate the data entry
work of Non Standard Expense.

[[Problem and Solution Statements](/Documentation//ProblemAndSolution.md)]

## Usage

#### Create Required Queue, assets in Orchestrator and fill into config File

[[Config File details](/Documentation//Config.md)]

- Clone this project and publish to Orchestrator using UiPath Studio
- Open UiPath Assistant, you will find new process named _VoltasNonStandardExpenseEntryBot_
  ![screenshot-preview](Documentation/Screenshots/Assistant.jpg)
- Open Process and enable PIP Mode , Click RUN.
  ![screenshot-preview](Documentation/Screenshots/Assistant-start.jpg)
- Wait for process to start.
- A popup form will open, enter CRM credentials and click on submit
  ![screenshot-preview](Documentation/Screenshots/userIdPasswordform.jpg)
- Again another popup form will open, now specify excel file path, and excel sheet name
  ![screenshot-preview](Documentation/Screenshots/excelFile.jpg)

- BOT will start Data Entry Process


## License

 [![AGPL License](https://img.shields.io/badge/LICENSE-GNU%20AGPLv3-brightgreen)](https://www.gnu.org/licenses/agpl-3.0.en.html)

 VoltasNonStandardExpenseEntryBot is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY, without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU Affero General Public License along with VoltasNonStandardExpenseEntryBot. If not, see https://www.gnu.org/licenses/agpl-3.0.en.html.
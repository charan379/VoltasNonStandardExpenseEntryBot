#### Config.xlsx

### BOT Config file 

##### Settings Sheet

| Name                     | Value                            | Description                                                                                                                                                                                        |
| ------------------------ | -------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| OrchestratorQueueName    | TadaVoltas                      | Orchestrator queue Name. The value must match with the queue name defined on Orchestrator.                                                                                                         |
| OrchestratorQueueFolder  | Shared/Non_stand                  | Folder name. The value must match a folder defined in Orchestrator and queue specified as OrchestratorQueueName should be created in this folder. For classic folders leave the value field empty. |
| logF_BusinessProcessName | VoltasNonStandardExpenseEntryBot | Logging field which allows grouping of log data of two or more subprocesses under the same business process name                                                                                   |


##### Constants Sheet

| Name                               | Value                                                             | Description                                                                                                                                                             |
| ---------------------------------- | ----------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| MaxRetryNumber                     | 3                                                                 | Must be 0 if working with Orchestrator queues. If > 0, the robot will retry the same transaction which failed with a system exception. Must be an integer value.        |
| MaxConsecutiveSystemExceptions     | 5                                                                 | The number of consecutive system exceptions allowed. If MaxConsecutiveSystemExceptions is reached, the job is stopped. To disable this feature, set the value to 0.     |
| ExScreenshotsFolderPath            | C:\\Users\\DEO\\Desktop\\VoltasNSESlogs\                                                  | Where to save exceptions screenshots - can be a full or a relative path.                                                                                                |
| LogMessage_GetTransactionData      | Processing Transaction Number:                                    | Static part of logging message. Calling Get Transaction Data.                                                                                                           |
| LogMessage_GetTransactionDataError | Error getting transaction data for Transaction Number:            | Static part of logging message. Error retrieving Transaction Data.                                                                                                      |
| LogMessage_Success                 | Transaction Successful.                                           | Static part of logging message. Processed Transaction succesful.                                                                                                        |
| LogMessage_BusinessRuleException   | Business rule exception.                                          | Static part of logging message. Processed Transaction failed with business exception.                                                                                   |
| LogMessage_ApplicationException    | System exception.                                                 | Static part of logging message. Processed Transaction failed with application exception.                                                                                |
| ExceptionMessage_ConsecutiveErrors | The maximum number of consecutive system exceptions was reached.  | Error message in case MaxConsecutiveSystemExceptions number is reached.                                                                                                 |
| RetryNumberGetTransactionItem      | 2                                                                 | The number of times Get Transaction Item activity is retried in case of an exception. Must be an integer >= 1.                                                          |
| RetryNumberSetTransactionStatus    | 2                                                                 | The number of times Set transaction status activity is retried in case of an exception. Must be an integer >= 1.                                                        |
| ShouldMarkJobAsFaulted             | FALSE                                                             | Must be TRUE or FALSE. If the value is TRUE and an error occurs in Initialization state or the MaxConsecutiveSystemExceptions is reached, the job is marked as Faulted. |
| outPutSrLogsFile                   | C:\\Users\\DEO\\Desktop\\VoltasNSESlogs\\OutPutSrLogsFile.xlsx |                                                                                                                                                                         |
| SrNotFoundStypeErr                 | SR Not Found                                           |                                                                                                                                                                         |


##### Assets Sheet

| Name   | Asset  | OrchestratorAssetFolder | Description (Assets will always overwrite other config) |
| ------ | ------ | ----------------------- | ------------------------------------------------------- |
| vl_url | vl_url | Shared/Voltas           | CRM URL                                                 |

Brief Synopsis

For various reasons, I was called upon to write a scheduler in powershell, specifically to run MS Access functions (from 2 specific Access 'databases') and Excel VBA functions from a number of Excel macro enabled files.

It was extended to run other powershell scripts, and could, quite easily, be further extended to run any external 'batch' type process.

Apart from the missing source functions (VBA and Powershell), the only other missing parts are the calendar and bank holiday tables (linked in via views).

In short, the entire process works by first generating the future schedule of tasks (for the next week, say) via stored proc uspTaskResultsAddNewTasksByDateRange then triggering the powershell script Scheduler_v1.ps1.

The primary weakness of the process is its lack of protection against a process timeout. This would be possible for a powershell task (not implemented) but not for a VBA task, as written.

Enjoy!

[Also, any feedback would be most welcome]

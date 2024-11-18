# SheetDataAccess

This library helps with CRUD operations with a Google Sheet in Google Apps Script projects.
Just copy the [/dist/SheetDataAccess.js](dist/SheetDataAccess.js) file into your project to get started.

## Using the library

Here's a basic example of using the lib:

```JavaScript
function getExampleTaskData() {
  /**
   * Assumes you have a Google Sheet with the following structure:
   * 
   * ____________________
   * | Sheet name: Task |
   * |-------------------
   * | id | task | done |
   * |-------------------
   * 
   */
  const ds = SheetDataAccess.create({ id: '<insert-your-sheet-id>' });

  const tasks = ds.collections.Task.data();

  console.log(tasks); 
  // Logs task objects in shape [{ id: 'task-id', task: 'task details': done: true/false }]

  const [taskOne] = tasks;
  task.done = true;

  // Saves the updated task back to the Sheet
  const updated = ds.collections.Task.updateOne(task);
}
```


For more examples, [check out Examples.ts](src/Examples.ts)
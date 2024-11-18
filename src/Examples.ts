type IProject = {
  id: string;
  name: string;
  priority: 'Low' | 'Medium' | 'High';
  due_date: Date;
} & SheetDataAccessRecordObjectKey;

type ITask = {
  id: string;
  project_id: string;
  title: string;
  notes: string;
  done: boolean;
  last_updated: Date;
} & SheetDataAccessRecordObjectKey;

type SheetModels = {
  Project: IProject;
  Task: ITask;
};

/**
 * Assumes you have a Google Sheet with the following structure:
 * 
 * ___________________________________
 * |            Project              |
 * |----------------------------------
 * | id | name | priority | due_date |
 * |----------------------------------
 * 
 * 
 * _________________________________________________________
 * |                         Task                          |
 * |--------------------------------------------------------
 * | id | project_id | title | notes | done | last_updated |
 * |--------------------------------------------------------
 * 
 */

/**
 * Helper to get the SheetDataAccess object
 */
const getSheetDataSource = () => {
  const id = '';
  return SheetDataAccess.create<SheetModels>({ id });
};

/**
 * Example of doing full text search on a sheet/table of data
 */
function example_fulltextsearch() {
  const ds = getSheetDataSource();
  const projects = ds.collections.Project.fts({ q: 'TypeScript', matchCell: true });

  console.log(projects);
}

/**
 * Example of getting related data in another Sheet for a given record
 */
function example_get_related_data() {
  const ds = getSheetDataSource();

  const project = ds.collections.Project.find('proj-001', 'id');
  if (!project) throw new Error('Could not find proj-001!');

  const itemsByRelated = ds.collections.Task.related('project_id')[project.id] || [];

  //  -- OR --

  const itemsByFilter = ds.collections.Task.data().filter(task => task.project_id === project.id);

  console.log(itemsByRelated)
  console.log(itemsByFilter);
}

/**
 * Example of updating a record in a Sheet
 */
function example_update_data() {
  const ds = getSheetDataSource();

  const task = ds.collections.Task.lookup('task-003', 'id');
  if (!task) throw new Error('Could not find task-001!');

  task.done = true;
  task.last_updated = new Date();

  const updated = ds.collections.Task.updateOne(task);

  console.log(updated);
}
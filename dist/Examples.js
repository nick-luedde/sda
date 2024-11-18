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
    return SheetDataAccess.create({ id });
};
/**
 * Example of doing full text search on a sheet/table of data
 */
function example_fulltextsearch() {
    const ds = getSheetDataSource();
    const projects = ds.collections.Project.fts({ q: 'TypeScript' });
    console.log(projects);
}
/**
 * Example of getting related data in another Sheet for a given record
 */
function example_get_related_data() {
    const ds = getSheetDataSource();
    const project = ds.collections.Project.find('proj-001', 'id');
    if (!project)
        throw new Error('Could not find proj-001!');
    const itemsByRelated = ds.collections.Task.related('project_id')[project.id] || [];
    //  -- OR --
    const itemsByFilter = ds.collections.Task.data().filter(task => task.project_id === project.id);
    console.log(itemsByRelated);
    console.log(itemsByFilter);
}
/**
 * Example of updating a record in a Sheet
 */
function example_update_data() {
    const ds = getSheetDataSource();
    const task = ds.collections.Task.lookup('task-003', 'id');
    if (!task)
        throw new Error('Could not find task-001!');
    task.done = true;
    task.last_updated = new Date();
    const updated = ds.collections.Task.updateOne(task);
    console.log(updated);
}
/**
 * Example of deleting a row
 */
function example_delete_row() {
    const ds = getSheetDataSource();
    const added = ds.collections.Task.addOne({
        id: 'task-1000',
        title: 'New to delete',
        done: true,
        last_updated: new Date(),
        project_id: '',
        notes: 'Some notes',
    });
    ds.collections.Task.delete([added]);
}
/**
 * Example of wiping a Sheet
 */
function example_wipe() {
    const ds = getSheetDataSource();
    const projects = ds.collections.Project.data();
    ds.collections.Project.wipe();
    ds.collections.Project.add(projects);
}
/**
 * Example of inspecting a Spreadsheet
 */
function example_inspect() {
    const ds = getSheetDataSource();
    const results = ds.inspect();
    console.log(results);
}
/**
 * Example of sorting the source sheet
 * WARNING: This could cook other peoples cached data, setting _keys out of sync
 */
function example_sort() {
    const ds = getSheetDataSource();
    ds.collections.Task.sort('last_updated');
    ds.collections.Task.sort('last_updated', true);
}
/**
 * Example of enforcing uniqueness on a column value
 */
function example_enforce_unique() {
    const ds = getSheetDataSource();
    const taskWithUniqueId = {
        id: 'task-unique-1000',
        project_id: 'none',
        title: 'Unique',
        notes: '',
        done: true,
        last_updated: new Date()
    };
    ds.collections.Task.enforceUnique(taskWithUniqueId, 'id'); // No error
    const taskWithNonUniqueId = {
        id: 'task-001',
        project_id: 'duplicate',
        title: 'duplicate',
        notes: 'duplicate',
        done: false,
        last_updated: new Date()
    };
    try {
        ds.collections.Task.enforceUnique(taskWithNonUniqueId, 'id'); // Error
    }
    catch (err) {
        console.error(err);
    }
}

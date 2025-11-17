const DEFAULT_FILTER = 'today';
const FILTER_LABELS = {
  today: '今日',
  overdue: '期限切れ',
  tomorrow: '明日',
  thisWeek: '今週',
  next7: '次の7日間',
  planned: '計画',
  noDue: '期限なし',
  completed: '完了済'
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Google Tasks ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getTaskLists() {
  return listTaskLists().map(function (list) {
    return { id: list.id, title: list.title };
  });
}

function getDashboardData(context) {
  var tz = Session.getScriptTimeZone();
  context = normalizeContext(context);
  var lists = listTaskLists();
  var listLookup = lists.reduce(function (map, list) {
    map[list.id] = list;
    return map;
  }, {});
  var targetLists = context.type === 'project' ? [listLookup[context.listId]] : lists;

  if (context.type === 'project' && !targetLists[0]) {
    throw new Error('指定したリストが見つかりませんでした。');
  }

  var tasks = [];
  targetLists.forEach(function (list) {
    fetchTasksForList(list.id).forEach(function (item) {
      var normalized = normalizeTask(item, list, tz);
      if (normalized) {
        tasks.push(normalized);
      }
    });
  });

  var filtered = applyFilter(tasks, context, tz);
  return {
    context: context,
    title: getContextTitle(context, listLookup),
    open: filtered.open,
    completed: filtered.completed,
    totals: filtered.totals,
    generatedAt: new Date().toISOString()
  };
}

function createTaskList(payload) {
  var title = payload && payload.title ? payload.title.trim() : '';
  if (!title) {
    throw new Error('プロジェクト名を入力してください。');
  }
  var list = Tasks.Tasklists.insert({ title: title });
  return { id: list.id, title: list.title };
}

function updateTaskList(payload) {
  var listId = payload && payload.id;
  var title = payload && payload.title ? payload.title.trim() : '';
  if (!listId || !title) {
    throw new Error('プロジェクト ID と名称は必須です。');
  }
  var list = Tasks.Tasklists.patch({ title: title }, listId);
  return { id: list.id, title: list.title };
}

function deleteTaskList(payload) {
  var listId = payload && payload.id;
  if (!listId) {
    throw new Error('プロジェクト ID は必須です。');
  }
  Tasks.Tasklists.remove(listId);
  return { success: true };
}

function createTask(payload) {
  var listId = payload && payload.listId;
  if (!listId) {
    throw new Error('listId は必須です。');
  }
  var resource = buildTaskResource(payload);
  var task = Tasks.Tasks.insert(resource, listId);
  return { id: task.id, listId: listId };
}

function updateTask(payload) {
  var targetListId = payload && payload.listId;
  var sourceListId = payload && (payload.sourceListId || targetListId);
  var taskId = payload && payload.id;
  if (!targetListId || !taskId || !sourceListId) {
    throw new Error('listId と task id は必須です。');
  }
  var resource = buildTaskResource(payload);
  if (sourceListId === targetListId) {
    var task = Tasks.Tasks.patch(resource, targetListId, taskId);
    return { id: task.id, listId: targetListId };
  }
  var existing = Tasks.Tasks.get(sourceListId, taskId);
  var merged = mergeTaskResource(existing, resource);
  var newTask = Tasks.Tasks.insert(merged, targetListId);
  Tasks.Tasks.remove(sourceListId, taskId);
  return { id: newTask.id, listId: targetListId };
}

function deleteTask(payload) {
  var listId = payload && payload.listId;
  var taskId = payload && payload.id;
  if (!listId || !taskId) {
    throw new Error('listId と task id は必須です。');
  }
  Tasks.Tasks.remove(listId, taskId);
  return { success: true };
}

function toggleTaskCompletion(payload) {
  var listId = payload && payload.listId;
  var taskId = payload && payload.id;
  if (!listId || !taskId) {
    throw new Error('listId と task id は必須です。');
  }
  var makeCompleted = !!payload.completed;
  var resource = {
    status: makeCompleted ? 'completed' : 'needsAction',
    completed: makeCompleted ? new Date().toISOString() : null
  };
  Tasks.Tasks.patch(resource, listId, taskId);
  return { success: true };
}

function normalizeContext(context) {
  var normalized = context || {};
  if (normalized.listId) {
    normalized.type = 'project';
  } else {
    normalized.type = 'filter';
  }
  if (normalized.type === 'filter' && !normalized.id) {
    normalized.id = DEFAULT_FILTER;
  }
  return normalized;
}

function listTaskLists() {
  var lists = [];
  var pageToken;
  do {
    var response = Tasks.Tasklists.list({
      maxResults: 100,
      pageToken: pageToken
    });
    if (response.items) {
      lists = lists.concat(response.items);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  lists.sort(function (a, b) {
    return a.title.localeCompare(b.title);
  });
  return lists;
}

function fetchTasksForList(listId) {
  var items = [];
  var pageToken;
  do {
    var response = Tasks.Tasks.list(listId, {
      showCompleted: true,
      showHidden: true,
      maxResults: 100,
      pageToken: pageToken
    });
    if (response.items) {
      response.items.forEach(function (item) {
        if (!item.deleted) {
          items.push(item);
        }
      });
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  return items;
}

function normalizeTask(item, list, tz) {
  if (!item) {
    return null;
  }
  var dueKey = item.due ? parseInt(Utilities.formatDate(new Date(item.due), tz, 'yyyyMMdd'), 10) : null;
  return {
    id: item.id,
    listId: list.id,
    listTitle: list.title,
    title: item.title || '無題のタスク',
    notes: item.notes || '',
    due: item.due || null,
    dueKey: dueKey,
    dueLabel: item.due ? buildDueLabel(item.due, tz) : '期限なし',
    status: item.status,
    isCompleted: item.status === 'completed',
    completedAt: item.completed || null,
    updatedAt: item.updated || null
  };
}

function buildDueLabel(dateString, tz) {
  var date = new Date(dateString);
  var label = Utilities.formatDate(date, tz, 'M月d日');
  var dayNames = ['月', '火', '水', '木', '金', '土', '日'];
  var index = getWeekdayIndex(date, tz);
  return label + ' (' + dayNames[(index - 1) % dayNames.length] + ')';
}

function sortTasksByDue(tasks) {
  return tasks.slice().sort(function (a, b) {
    var aNoDue = a.dueKey == null;
    var bNoDue = b.dueKey == null;
    if (aNoDue && bNoDue) {
      return a.title.localeCompare(b.title);
    }
    if (aNoDue) return -1;
    if (bNoDue) return 1;
    if (a.dueKey !== b.dueKey) {
      return a.dueKey - b.dueKey;
    }
    return a.title.localeCompare(b.title);
  });
}

function applyFilter(tasks, context, tz) {
  var filterId = context.type === 'project' ? 'project' : context.id;
  var now = new Date();
  var todayKey = getDateKey(now, tz);
  var tomorrowKey = getDateKey(addDays(now, 1), tz);
  var weekdayIndex = getWeekdayIndex(now, tz);
  var weekStart = addDays(now, -(weekdayIndex - 1));
  var weekEnd = addDays(weekStart, 6);
  var nextSevenEndKey = getDateKey(addDays(now, 7), tz);
  var weekStartKey = getDateKey(weekStart, tz);
  var weekEndKey = getDateKey(weekEnd, tz);
  var filtered = tasks;

  switch (filterId) {
    case 'today':
      filtered = tasks.filter(function (task) {
        return !task.isCompleted && task.dueKey === todayKey;
      });
      break;
    case 'overdue':
      filtered = tasks.filter(function (task) {
        return !task.isCompleted && task.dueKey && task.dueKey < todayKey;
      });
      break;
    case 'tomorrow':
      filtered = tasks.filter(function (task) {
        return !task.isCompleted && task.dueKey === tomorrowKey;
      });
      break;
    case 'thisWeek':
      filtered = tasks.filter(function (task) {
        return !task.isCompleted && task.dueKey && task.dueKey >= weekStartKey && task.dueKey <= weekEndKey;
      });
      break;
    case 'next7':
      filtered = tasks.filter(function (task) {
        return !task.isCompleted && task.dueKey && task.dueKey >= todayKey && task.dueKey <= nextSevenEndKey;
      });
      break;
    case 'planned':
      filtered = tasks.filter(function (task) {
        return !task.isCompleted && task.dueKey;
      });
      break;
    case 'noDue':
      filtered = tasks.filter(function (task) {
        return !task.isCompleted && !task.dueKey;
      });
      break;
    case 'completed':
      filtered = tasks.filter(function (task) {
        return task.isCompleted;
      });
      break;
    case 'project':
      filtered = tasks;
      break;
    default:
      filtered = tasks.filter(function (task) {
        return !task.isCompleted;
      });
      break;
  }

  var ordered = sortTasksByDue(filtered);

  var open = ordered.filter(function (task) {
    return !task.isCompleted;
  });
  var completed = ordered.filter(function (task) {
    return task.isCompleted;
  });

  return {
    open: open,
    completed: completed,
    totals: {
      open: open.length,
      completed: completed.length
    }
  };
}

function getContextTitle(context, listLookup) {
  if (context.type === 'project') {
    var list = listLookup[context.listId];
    return list ? list.title + ' のタスク' : 'リストのタスク';
  }
  return (FILTER_LABELS[context.id] || 'タスク') + 'の一覧';
}

function buildTaskResource(payload) {
  var resource = {};
  if (payload.title != null) {
    resource.title = payload.title;
  }
  if (payload.notes != null) {
    resource.notes = payload.notes;
  }
  if (Object.prototype.hasOwnProperty.call(payload, 'dueDate')) {
    resource.due = payload.dueDate ? buildDueDateString(payload.dueDate) : null;
  }
  return resource;
}

function mergeTaskResource(existing, overrides) {
  var resource = {
    title: existing && existing.title ? existing.title : '無題のタスク',
    notes: existing && existing.notes ? existing.notes : '',
    due: existing && existing.due ? existing.due : null,
    status: existing && existing.status ? existing.status : 'needsAction',
    completed: existing && existing.completed ? existing.completed : null
  };

  ['title', 'notes', 'due'].forEach(function (field) {
    if (Object.prototype.hasOwnProperty.call(overrides, field)) {
      resource[field] = overrides[field];
    }
  });

  if (resource.status !== 'completed') {
    resource.completed = null;
  } else if (!resource.completed) {
    resource.completed = new Date().toISOString();
  }

  return resource;
}

function buildDueDateString(value) {
  if (!value) {
    return null;
  }
  if (value.indexOf('T') > -1) {
    return value;
  }
  return value + 'T00:00:00.000Z';
}

function addDays(date, days) {
  var copy = new Date(date.getTime());
  copy.setDate(copy.getDate() + days);
  return copy;
}

function getDateKey(date, tz) {
  return parseInt(Utilities.formatDate(date, tz, 'yyyyMMdd'), 10);
}

function getWeekdayIndex(date, tz) {
  var code = Utilities.formatDate(date, tz, 'E');
  var map = { Sun: 7, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };
  return map[code] || 1;
}

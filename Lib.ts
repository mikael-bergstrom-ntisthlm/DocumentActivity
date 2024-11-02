function GetHistory(docUrl: string, userResourceName: string): Date[] {

  // docUrl = "https://docs.google.com/document/d/1ug7QXhC9t6BDSHoGmjC9kwN9612gvd_j7z5RVG-bXIo/edit";
  // userEmail = "abdullahi.abdirahman@elev.ga.ntig.se";

  let editTimestamps: Date[] = [];
  // let editWeeks: Set<string> = new Set();

  const document = DocumentApp.openByUrl(docUrl);

  // Get all edits
  let result = DriveActivity.Activity?.query({
    "ancestorName": "items/" + document.getId(),
    "filter": "detail.action_detail_case:EDIT"
  });

  if (!result) return [];

  // Go through all edits
  result.activities?.forEach(activity => {
    if (!activity.actors) return;

    // Go through all editors
    activity.actors.forEach(actor => {
      if (!actor.user?.knownUser?.personName
        || !activity.timestamp
      ) return;
      
      if (actor.user.knownUser.personName === userResourceName) {
        // let weekNum = Utilities.formatDate(new Date(activity.timestamp), Session.getScriptTimeZone(), "w");
        // editWeeks.add(weekNum);
        // Logger.log(weekNum);
        editTimestamps.push(new Date(activity.timestamp));
      }
    })
  })

  return editTimestamps;
}

// function GetEmailOfUser(userId: string): string
// {
//   let person = People.People?.get(userId,
//     {
//       personFields: 'emailAddresses'
//     }
//   );

//   if (!person?.emailAddresses
//     || person.emailAddresses.length < 1
//     || person.emailAddresses[0].value === undefined
//   ) return "nonefound@email.com";

//   return person.emailAddresses[0].value;
// }

function GetUserResourceName(userQuery: string)
{
  let persons = People.People?.searchDirectoryPeople({
    query: userQuery,
    readMask: 'names',
    sources: [2]
  });

  if (!persons || !persons.people || persons.people.length < 1 || !persons.people[0].resourceName) return "";

  return persons.people[0].resourceName;
}
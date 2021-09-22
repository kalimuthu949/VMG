import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
//import * as $ from "jquery";
import {sp} from "@pnp/pnpjs";
import { Web } from "@pnp/sp/webs";

import * as strings from 'ApprovalCommandsetCommandSetStrings';

var arrSelectedRows=[];

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IApprovalCommandsetCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
  ApprovalFolderURL:string;
  RejectedFolderURL:string;
  FileName:string;
  FileURL:string;
}

const LOG_SOURCE: string = 'ApprovalCommandsetCommandSet';

var listweb = Web("https://veritasmg.sharepoint.com/sites/VMGsProjectManagementTeam"); 

export default class ApprovalCommandsetCommandSet extends BaseListViewCommandSet<IApprovalCommandsetCommandSetProperties> {

  @override
  public onInit(): Promise<void> 
  {
    Log.info(LOG_SOURCE, 'Initialized ApprovalCommandsetCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void 
  {
    var Libraryurl = this.context.pageContext.list.title;

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    const compareOneCommand2: Command = this.tryGetCommand('COMMAND_2');

      var HoldFolderName="Screened Resumes";
      var ApprovalFolderName="Approved Resumes";
      var RejectedFolderName="Rejected Resumes";
      var ApprovalFolderURL="";
      var RejectedFolderURL="";
      var FileName="";
      var FileURL=""
      var Libraryname="Documents";
      var CandidateCode="";
      var ItemId="";

      var flgDoc=true;//check whether it is document or not..
      var flgRootFolder=false;//check whether it is root folder or not..

      if(event.selectedRows.length>0)
      {
        arrSelectedRows=[];
        for(var i=0;i<event.selectedRows.length;i++)
        {var jCode ="" 
            event.selectedRows[i]["_values"].forEach(async function(val,key)
            {
              if(key=="ContentType"&&val=="Folder")
              flgDoc=false;

              if(key=="FileRef")
              {
                FileURL=val;
                

                var RootFolderUrl=val.split("/");
                var RootFolderName=RootFolderUrl[RootFolderUrl.length-2];
                FileName=RootFolderUrl[RootFolderUrl.length-1];

                if(HoldFolderName.toLowerCase()==RootFolderName.toLowerCase())
                {
                  flgRootFolder=true;

                  ApprovalFolderURL=RootFolderUrl.slice(0, -2).join("/")+"/"+ApprovalFolderName;
                  RejectedFolderURL=RootFolderUrl.slice(0, -2).join("/")+"/"+RejectedFolderName;
                  
                  jCode = (val.split('1. Job Code/')[1]).split('/')[0]
                }
              }

              if(key=="CandidateCode"){
                CandidateCode=val;
                }
                if(key=="UniqueId"){
                  ItemId=val;
                  }

              await [];
              
            });

            arrSelectedRows.push({"ApprovalFolderURL":ApprovalFolderURL,"RejectedFolderURL":RejectedFolderURL,"FileName":FileName,"FileURL":FileURL,"CandidateCode":CandidateCode,"ItemId":ItemId,JobCode:jCode});
        }
      }

      /*this.properties.ApprovalFolderURL=ApprovalFolderURL;
      this.properties.RejectedFolderURL=RejectedFolderURL;
      this.properties.FileName=FileName;
      this.properties.FileURL=FileURL;*/

    if (compareOneCommand) 
    {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length > 0&&flgDoc&&flgRootFolder&&Libraryurl==Libraryname;

    }

    if (compareOneCommand2) 
    {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand2.visible = event.selectedRows.length > 0&&flgDoc&&flgRootFolder&&Libraryurl==Libraryname;
    }

    

  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void 
  {
    switch (event.itemId) {
      case 'COMMAND_1':
        //updateApprovalFolder(this.properties.ApprovalFolderURL,this.properties.FileName,this.properties.FileURL);
        updateApprovalFolder();
        break;
      case 'COMMAND_2':
        //updateRejectFolder(this.properties.RejectedFolderURL,this.properties.FileName,this.properties.FileURL);
        updateRejectFolder();
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}

//async function updateApprovalFolder(ApprovalFolder,FileNametomove,sourceURL)
async function updateApprovalFolder()
{    
    var count=0;
    for(var i=0;i<arrSelectedRows.length;i++)
    {
      if(arrSelectedRows[i].ApprovalFolderURL&&arrSelectedRows[i].FileName)
    {
        // destination is a server-relative url of a new file
        const destinationUrl = arrSelectedRows[i].ApprovalFolderURL+"/"+arrSelectedRows[i].FileName;
       const resumeFolderUrl = "/sites/HRPrivateNew/Shared%20Documents/1. Recruiting & Onboarding/2. Resumes/"+arrSelectedRows[i].JobCode+"/"+arrSelectedRows[i].FileName
        // await sp.web.getFileByServerRelativePath(arrSelectedRows[i].FileURL).moveTo(destinationUrl).then(async function(data)
        await sp.web.getFileById(arrSelectedRows[i].ItemId).copyTo(resumeFolderUrl, true).then(async()=>{
          await sp.web.getFileById(arrSelectedRows[i].ItemId).moveTo(destinationUrl).then(async function(data)
          {
            count++;
  
            var code=arrSelectedRows[i].CandidateCode;
            var ResumeURL="";
            var Itemid;
        
            await listweb.lists.getByTitle("VMG Recruitment Tracker")
            .items.select("*").filter("CandidateCode eq '"+code+"'").get()
            .then(async (item)=>{
              console.log(item[0].ID);
              Itemid=item[0].ID;
              var Url=item[0].ResumeURL;
              ResumeURL=Url.split('Screened Resumes').join('Approved Resumes');
            }).catch((error)=>
            {
              Dialog.alert(`someting went wrong.please try again`).then(function(){
                location.reload();
                return false;
              });
            });
        
              await listweb.lists
                  .getByTitle("VMG Recruitment Tracker")
                  .items.getById(Itemid).update({"ResumeURL":ResumeURL})
                  .then(async function (data) 
                  {
        
                  }).catch((error)=>
                  {
                    //console.log(error);
                    Dialog.alert(`someting went wrong.please try again`).then(function(){
                      location.reload();
                      return false;
                    });
                  });
  
          }).catch(function(error)
          {
              Dialog.alert(`someting went wrong.please try again`).then(function(){
                location.reload();
                return false;
              });
          }).catch(function(error)
          {
              Dialog.alert(`someting went wrong.please try again`).then(function(){
                location.reload();
                return false;
              });
          })  
        })

    }
    }
            Dialog.alert(`Approved Sucessfully`).then(function(){
              location.reload();
            });

    
}

//async function updateRejectFolder(ApprovalFolder,FileNametomove,sourceURL)
async function updateRejectFolder(

)
{
    /*await sp.web.lists.getByTitle("Projects").items.add({"Title":"Test"}).then(function(data)
    {
        Dialog.alert(`Approve Sucessfully`);

    }).catch(function(error)
    {
        Dialog.alert(`someting went wrong.please try again`);
    })*/

    /*if(ApprovalFolder&&FileNametomove)
    {
        // destination is a server-relative url of a new file
        const destinationUrl = ApprovalFolder+"/"+FileNametomove;

        await sp.web.getFileByServerRelativePath(sourceURL).moveTo(destinationUrl).then(function(data)
        {
            Dialog.alert(`Rejected Sucessfully`).then(function(){
              location.reload();
            });

        }).catch(function(error)
        {
            Dialog.alert(`someting went wrong.please try again`).then(function(){
              location.reload();
            });
        })
    }*/


    // var count=0;
    // for(var i=0;i<arrSelectedRows.length;i++)
    // {
    //   if(arrSelectedRows[i].RejectedFolderURL&&arrSelectedRows[i].FileName)
    // {
    //     // destination is a server-relative url of a new file
    //     const destinationUrl = arrSelectedRows[i].RejectedFolderURL+"/"+arrSelectedRows[i].FileName;

    //     await sp.web.getFileByServerRelativePath(arrSelectedRows[i].FileURL).moveTo(destinationUrl).then(async function(data)
    //     {
    //       count++;

    //       const destinationUrl = ApprovalFolder+"/"+FileNametomove;
    //       await sp.web.getFileByServerRelativePath(sourceURL).moveTo(destinationUrl).then(function(data)
    //       {
    //           Dialog.alert(`Rejected Sucessfully`).then(function(){
    //             location.reload();
    //           });
    //       }).catch(function(error)
    //       {
    //           Dialog.alert(`someting went wrong.please try again`).then(function(){
    //             location.reload();
    //           });
    //       })
    //   }*/
      var count=0;
      for(var i=0;i<arrSelectedRows.length;i++)
      {
        if(arrSelectedRows[i].RejectedFolderURL&&arrSelectedRows[i].FileName)
      {
          // destination is a server-relative url of a new file
          const destinationUrl = arrSelectedRows[i].RejectedFolderURL+"/"+arrSelectedRows[i].FileName;
          //await sp.web.getFileByServerRelativeUrl(arrSelectedRows[i].FileURL).moveTo(destinationUrl).then(async function(data)
          await sp.web.getFileById(arrSelectedRows[i].ItemId).moveTo(destinationUrl).then(async function(data)
          {
            count++;
            var code=arrSelectedRows[i].CandidateCode;
            var ResumeURL="";
            var Itemid;
        
            await listweb.lists.getByTitle("VMG Recruitment Tracker")
            .items.select("*").filter("CandidateCode eq '"+code+"'").get()
            .then(async (item)=>{
              console.log(item[0].ID);
              Itemid=item[0].ID;
              var Url=item[0].ResumeURL;
              //Url.replace('Screened','Approved');
              ResumeURL=Url.split('Screened Resumes').join('Rejected Resumes');
            }).catch((error)=>
            {
              console.log(error);
            });
        
              await listweb.lists
                  .getByTitle("VMG Recruitment Tracker")
                  .items.getById(Itemid).update({"ResumeURL":ResumeURL})
                  .then(async function (data) 
                  {
        
                  }).catch((error)=>
                  {
                    console.log(error);
                  });

        }).catch(function(error)
        {
            Dialog.alert(`someting went wrong.please try again`).then(function(){
              location.reload();
              return false;
            });
        })
    }
    }
            Dialog.alert(`Rejected Sucessfully`).then(function(){
              location.reload();
            });

    
}

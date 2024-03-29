import * as React from "react";
import TreeView from "@material-ui/lab/TreeView";
import ExpandMoreIcon from "@material-ui/icons/ExpandMore";
import ChevronRightIcon from "@material-ui/icons/ChevronRight";
import TreeItem from "@material-ui/lab/TreeItem";
import { graph } from "@pnp/graph/presets/all";
import "../assets/Css/NewPivot.css";
import { MSGraphClient } from "@microsoft/sp-http";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { makeStyles } from "@material-ui/core/styles";
let allUserData = [];

import SPServices from "./SPServices";

const useStyles = makeStyles({
  label: {
    backgroundColor: "#fff",
    // color: "red"
  },
});
export default function NewPivot(props) {
  const classes = useStyles(props);

  const [peopleList, setPeopleList] = React.useState([]);
  const [department, setdepartment] = React.useState([]);
  const [designationdetails, setdesignationdetails] = React.useState([]);
  const [loader, setloader] = React.useState(false);

  React.useEffect(function () {
    setloader(true);
    getEmployeeDetails();
    // getallusers();
  }, []);

  function removeDuplicates(arr) {
    return arr.filter((item, index) => arr.indexOf(item) === index);
  }

  const getEmployeeDetails = () => {
    SPServices.SPReadItems({
      Listname: "EmployeeGroupDetails",
      Select: "*,Manager/Title,Manager/Id,Manager/EMail",
      Expand: "Manager",
    })
      .then((data: any) => {
        let employeeArr = [];
        for (const item of data) {
          employeeArr.push({
            mail: item.Email,
            id: item.Title,
            displayName: [item.FirstName, item.LastName].join(" "),
            userPrincipalName: item.UserPrincipalName,
            jobTitle: item.JobTitle,
            givenName: item.FirstName,
            surname: item.LastName,
            businessPhones: item.PhoneNumber ? item.PhoneNumber.split(",") : [],
            department: item.Department,
            officeLocation: item.Zone,
            manager: item.ManagerId ? item.Manager.Title : "",
          });
        }

        // employeeArr.sort((a, b) => sortFunction(a, b, "displayName"));
        bindData(employeeArr);
      })
      .catch((error) => {
        console.log(error);
        setloader(false);
      });
  };

  function bindData(data) {
    //let devDomain = "chandrudemo.onmicrosoft.com";
    // let devDomain = "hosthealthcare.onmicrosoft.com";
    const users = [];
    let depts = [];
    for (let i = 0; i < data.length; i++) 
    {
      
      /*Changes start */
      // let userIdentity = data[i].identities[0].issuer;
      // let userPrinName = data[i].userPrincipalName
      //   ? data[i].userPrincipalName
      //   : "";

      // if (!props.propertyToggle) {
      //   if (
      //     userIdentity.toLowerCase() == devDomain &&
      //     !userPrinName.includes("#EXT#")
      //   ) {
      //     users.push({
      //       imageUrl:
      //         "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
      //       isValid: true,
      //       Email: data[i].mail,
      //       ID: data[i].id,
      //       key: i,
      //       text: data[i].displayName,
      //       jobTitle: data[i].jobTitle,
      //       mobilePhone: data[i].mobilePhone,
      //       department: data[i].department,
      //     });

      //     if (data[i].department) depts.push(data[i].department.trim());
      //   }
      // } else {
      //   users.push({
      //     imageUrl:
      //       "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
      //     isValid: true,
      //     Email: data[i].mail,
      //     ID: data[i].id,
      //     key: i,
      //     text: data[i].displayName,
      //     jobTitle: data[i].jobTitle,
      //     mobilePhone: data[i].mobilePhone,
      //     department: data[i].department,
      //   });

      //   if (data[i].department) depts.push(data[i].department.trim());
      // }
      /* changes end */

      users.push({
        imageUrl: "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
        isValid: true,
        Email: data[i].mail,
        ID: data[i].id,
        key: i,
        text: data[i].displayName,
        jobTitle: data[i].jobTitle,
        // mobilePhone: data[i].mobilePhone,
        // mobilePhone:
        //   data[i].businessPhones.length > 0 ? data[i].businessPhones[0] : "",
        department: data[i].department,
      });
      if (data[i].department) depts.push(data[i].department.trim());
    }

    depts = removeDuplicates(depts);
    let designations = [];
    for (let i = 0; i < depts.length; i++) {
      designations.push({ Dept: depts[i], UserCount: 0, Designations: [] });

      for (let j = 0; j < users.length; j++) {
        if (users[j].department) {
          if (users[j].department.trim() == depts[i].trim()) {
            if (users[j].jobTitle) {
              let obj = "";
              if (designations[i].Designations.length > 0) {
                if (users[j].jobTitle)
                  obj = designations[i].Designations.find(
                    (o) => o.Designation.trim() == users[j].jobTitle.trim()
                  );
              }
              if (obj) {
                let index = designations[i].Designations.findIndex(
                  (o) => o.Designation.trim() == users[j].jobTitle.trim()
                );
                designations[i].UserCount = designations[i].UserCount + 1;
                designations[i].Designations[0].count =
                  designations[i].Designations[0].count + 1;
              } else {
                designations[i].UserCount = designations[i].UserCount + 1;
                designations[i].Designations.push({
                  Designation: users[j].jobTitle,
                  count: 1,
                });
              }
            }
          }
        }
      }
    }

    let sortedArray = designations.sort(function (a, b) {
      if (a.Dept < b.Dept) {
        return -1;
      }
      if (a.Dept > b.Dept) {
        return 1;
      }
      return 0;
    });
    //let sortedArray = designations;

    console.log(sortedArray);
    setdesignationdetails([...sortedArray]);
    setdepartment([...depts]);
    setPeopleList([...users]);
    setloader(false);
  }

  async function getNextusers(skiptoken) {
    //await graph.users.select("department,mail,id,displayName,jobTitle,mobilePhone").top(10)
    await props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("users")
          .select(
            "department,mail,id,displayName,jobTitle,mobilePhone,identities,userPrincipalName"
          )
          .top(999)
          .skipToken(skiptoken)
          .get()
          .then(function (data) {
            allUserData = [...allUserData, ...data.value];
            let strtoken = "";
            if (data["@odata.nextLink"]) {
              strtoken = data["@odata.nextLink"].split("skipToken=")[1];
              getNextusers(data["@odata.nextLink"].split("skipToken=")[1]);
            } else {
              bindData(allUserData);
            }
          })
          .catch(function (error) {
            console.log(error);
            setloader(false);
          });
      });
  }

  async function getallusers() {
    allUserData = [];
    //await graph.users.select("department,mail,id,displayName,jobTitle,mobilePhone").top(10)
    await props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("users")
          .select(
            "department,mail,id,displayName,jobTitle,mobilePhone,identities,userPrincipalName"
          )
          .top(999)
          .get()
          .then(function (data) {
            allUserData = data.value;
            let strtoken = "";
            if (data["@odata.nextLink"]) {
              strtoken = data["@odata.nextLink"].split("skiptoken=")[1];
              getNextusers(data["@odata.nextLink"].split("skiptoken=")[1]);
            } else {
              bindData(data.value);
            }
          })
          .catch(function (error) {
            console.log(error);
            setloader(false);
          });
      });
  }

  return (
    <div className="clsPivot">
      {loader ? (
        <div className="spinnerBackground">
          <Spinner className="clsSpinner" size={SpinnerSize.large} />
        </div>
      ) : (
        <></>
      )}
      <TreeView
        aria-label="file system navigator"
        defaultCollapseIcon={<ExpandMoreIcon />}
        defaultExpandIcon={<ChevronRightIcon />}
      >
        {designationdetails.map(function (item, index) {
          let count = item.Designations.length;
          let sortedArraydesignation = item.Designations.sort(function (a, b) {
            if (a.Designation < b.Designation) {
              return -1;
            }
            if (a.Designation > b.Designation) {
              return 1;
            }
            return 0;
          });
          return (
            <TreeItem
              nodeId={index.toString()}
              label={item.Dept + " (" + item.UserCount + ")"}
            >
              {sortedArraydesignation.map(function (item, index) {
                let labelvalue = item.Designation + " (" + item.count + ")";
                return (
                  <TreeItem
                    nodeId={designationdetails.length.toString()}
                    classes={{ label: classes.label }}
                    label={labelvalue}
                  />
                );
              })}
            </TreeItem>
          );
        })}
      </TreeView>
    </div>
  );
}

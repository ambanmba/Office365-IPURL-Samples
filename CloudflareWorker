/**
 * This will pull changes for Office 365 IP Addresses and URLs.
 * For more information please see http://aka.ms/ipurlws
 * Based on a script from https://github.com/pandrew1/Office365-IPURL-Samples
 * Modified to work directly as a Cloudflare Worker with no additional infrastructure
 * Modified to allow a search parameter to specify which version to start from
 * Modified to use a new ClientRequestID with each request
 * Version 1.0 - 26-Sep-2021 - By: Alastair Bor
 */


async function handleRequest(request) {
  const init = {
    headers: {
      "content-type": "application/json;charset=UTF-8",
    },
  }
  
  const msinit = {
    headers: {
      "Content-Type": "text/html",
    },
  }
  
  const host = "https://endpoints.office.com"
  const dir = "/changes/worldwide/"
  const uuid = uuidv4()
  const requiredVerRegExp = new RegExp(/^[0-9]{10}$/)

  var ver = "2018062700" // Default is to go back to 2018062700 (the first entry)
  var returnmessage = "<p>Use /?ver=YYYYMMDDNN in the URL to start from a specific version number. E.g. /?ver=2020061600</p>"
  
  const version = new URL(request.url)
  if (version.searchParams.has('ver')) {
     ver = version.searchParams.get('ver')
     if (requiredVerRegExp.test(ver)) {
        returnmessage = "<p>Changes shown from version: " + ver + "</p>"
        if (ver > 10) { // Microsoft only sends changes SINCE the version number, need to decrement by 1 to include the version number as well
            ver = ver - 1
        } 
     }
     else {
        ver = "2018062700"
        returnmessage = "<p>Invalid version, format must be YYYYMMDDNN (e.g. " + ver + ")</p>"
     }
  }

  const url = host + dir + ver + "?ClientRequestId=" + uuid

  const response = await fetch(url, init).then(response => response.json())
  const msresults = 
        "<!DOCTYPE html><html><head><title>Table of changes</title>" +
        "<style>table, th, td {border: 1px solid black;border-collapse: collapse;font-family: Segoe UI, Helvetica, sans-serif;}" +
        "th, td {padding: 8px;} p {font-family: Segoe UI, Helvetica, sans-serif;} h2 {font-family: Segoe UI, Helvetica, sans-serif;}" +
        "</style></head><body>" +
        "<p>Recent changes for Office 365 IP Addresses and URLs. For more information please see <a href='http://aka.ms/ipurlws'>http://aka.ms/ipurlws</a></p>" +
        returnmessage +
        "<p>ClientRequestID="+uuid+"</p><br>" +
        formatChanges(JSON.parse(JSON.stringify(response))) + "</body></html>"
 
  return new Response(msresults, msinit)
}

addEventListener("fetch", event => {
  return event.respondWith(handleRequest(event.request))
})

function uuidv4() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

function formatChanges(changes) {
    var out = ""
    changes.forEach(function(obj) { 
        out += "<h2>Change {0}</h2>".replace("{0}", obj.id)
        out += "<p>Endpoint Set ID: " + obj.endpointSetId 
        out += ", Disposition: " + obj.disposition 
        out += ", Version " + obj.version 
        if (isdef(obj.impact)) {
            out += ", Impact: " + obj.impact 
        }
        out += "</p><table style='width:100%'><tbody>"
        if (isdef(obj.previous)) {
            out += "<tr><td><b>Previous</b></td></tr>" + formatSub(obj.previous)
        } 
        if (isdef(obj.current)) {
            out += "<tr><td><b>Current</b></td></tr>" + formatSub(obj.current)
        } 
        if (isdef(obj.add)) {
            out += "<tr><td><b>Add</b></td></tr><tr><td>"
            if (isdef(obj.add.effectiveDate)) {
                out += "Effective Date: {0} ".replace("{0}", obj.add.effectiveDate)
            }
            if (isdef(obj.add.urls)) {
                out += "URLs: "
                var first = true
                obj.add.urls.forEach(function(url) {
                    if (!first) { out += ", " }
                    out += url
                    first = false
                })
                out += " "
            }
            if (isdef(obj.add.ips)) {
                out += "IP Addresses: "
                var first = true
                obj.add.ips.forEach(function(url) {
                    if (!first) { out += ", "}
                    out += url
                    first = false
                })
            }
            out += "</td></tr>"
        } else if (isdef(obj.remove)) {
            out += "<tr><td><b>Remove</b></td></tr><tr><td>"
            if (isdef(obj.remove.urls)) {
                out += "URLs: "
                var first = true
                obj.remove.urls.forEach(function(url) {
                    if (!first) { out += ", "}
                    out += url
                    first = false
                })
                out += " "
            }
            if (isdef(obj.remove.ips)) {
                out += "IP Addresses: "
                var first = true
                obj.remove.ips.forEach(function(url) {
                    if (!first) { out += ", "}
                    out += url
                    first = false
                })
            }
            out += "</td></tr>"
        } 
        out += "</tbody></table>"
    })
    return out
}

function formatSub(node){
    var out = "<tr><td>" 
    var first = true
    if (isdef(node.expressRoute)) {
        if (!first) { out += ", "} 
        first = false
        out += "ExpressRoute: " + node.expressRoute 
    }
    if (isdef(node.serviceArea)) {
        if (!first) { out += ", "} 
        first = false
        out += "serviceArea: " + node.serviceArea 
    }
    if (isdef(node.category)) {
        if (!first) { out += ", "} 
        first = false
        out += "category: " + node.category 
    }
    if (isdef(node.notes)) {
        if (!first) { out += ", "} 
        first = false
        out += "notes: " + node.notes 
    }
    if (isdef(node.required)) {
        if (!first) { out += ", "} 
        first = false
        out += "required: " + node.required 
    }
    if (isdef(node.tcpPorts)) {
        if (!first) { out += ", "} 
        first = false
        out += "tcpPorts: " + node.tcpPorts 
    }
    if (isdef(node.udpPorts)) {
        if (!first) { out += ", "} 
        first = false
        out += "udpPorts: " + node.udpPorts 
    }
    out += "</td></tr>"
    return out
}

function isdef(node){
    return (typeof node !== 'undefined')
}

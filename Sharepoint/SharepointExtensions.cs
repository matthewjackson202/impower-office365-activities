using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using Newtonsoft.Json;

namespace Impower.Office365.Sharepoint
{
    public static class SharepointExtensions
    {
        //TODO: Make this better?
        public static string GetSharepointHostNameFromUrl(string url)
        {
            return url.Replace("http://", String.Empty).Replace("https://", String.Empty).Split('/')[0];
        }
        //TODO: Make this better?
        public static string GetSharepointSitePathFromUrl(string url)
        {
            int index = url.IndexOf("/sites/");
            if (index < 0)
            {
                throw new Exception("Could not find site path from URL");
            }
            return url.Substring(index).TrimEnd('/');
        }
        public static async Task<Site> GetSharepointSite(
            this GraphServiceClient client,
            CancellationToken token,
            string webUrl
        )
        {
            var hostName = GetSharepointHostNameFromUrl(webUrl);
            var sitePath = GetSharepointSitePathFromUrl(webUrl);
            Console.WriteLine(hostName + " - " + sitePath);

            return await client.Sites.GetByPath(sitePath, hostName).Request().GetAsync(token);
        }
        public static async Task<Drive> GetSharepointDrive(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveName
        )
        {
            var site = await client.Sites[siteId].Request().GetAsync(token);
            var allDrives = await client.Sites[siteId].Drives.Request().GetAsync(token);
            var matchingDrives = allDrives.Where(drive => drive.Name == driveName);
            if (matchingDrives.Any())
            {
                return matchingDrives.First();
            }
            else
            {
                throw new Exception("Cannot find matching drive.");
            }
        }
        public static async Task<DriveItem> GetSharepointDriveItem(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string itemId
        )
        {
            return await client.Sites[siteId].Drives[driveId].Items[itemId].Request().Expand(item => item.ListItem).GetAsync(token);
        }
        public static async Task<List> GetSharepointList(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string listId
        )
        {
            return await client.Sites[siteId].Lists[listId].Request().GetAsync(token);

        }
        public static async Task<FieldValueSet> UpdateSharepointDriveItemFields(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string itemId,
            FieldValueSet fieldValueSet
        )
        {

            return await client.Sites[siteId].Drives[driveId].Items[itemId].ListItem.Fields.Request().UpdateAsync(fieldValueSet,token);

        }
        public static async Task<ListItem> GetSharepointListItem(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string listId,
            string itemId
        )
        {
            return await client.Sites[siteId].Lists[listId].Items[itemId].Request().GetAsync(token);
        }
        public static async Task<List<DriveItem>> GetSharepointDriveItemsByPath(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string path
        )
        {
            IDriveItemChildrenCollectionRequest request;
            string folder = Path.GetDirectoryName(path);
            string filename = Path.GetFileName(path);
            Console.WriteLine(folder + " - " + filename);
            var requestBase = client.Sites[siteId].Drives[driveId].Root;
            
            if (String.IsNullOrWhiteSpace(folder))
            {
                request = requestBase.Children.Request();
            }
            else
            {
                request = requestBase.ItemWithPath(folder).Children.Request();
            }

            var items = await request.GetAsync(token);
            if (String.IsNullOrWhiteSpace(filename))
            {
                return items.ToList();
            }
            else
            {
                return items.Where(item => item.Name == filename).ToList();
            }

        }
    }
}

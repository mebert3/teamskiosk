// Make sure to install the necessary dependencies
const { CallClient, VideoStreamRenderer, LocalVideoStream } = require('@azure/communication-calling');
const { AzureCommunicationTokenCredential } = require('@azure/communication-common');
const { AzureLogger, setLogLevel } = require("@azure/logger");
// Set the log level and output
setLogLevel('verbose');
AzureLogger.log = (...args) => {
    console.log(...args);
};
// Calling web sdk objects
let teamsCallAgent;
let deviceManager;
let call;
let incomingCall;
let localVideoStream;
let localVideoStreamRenderer;
// UI widgets
let userAccessToken = document.getElementById('user-access-token');
let calleeTeamsUserId = document.getElementById('callee-teams-user-id');
let initializeCallAgentButton = document.getElementById('initialize-teams-call-agent');
let mainframe = document.getElementById('mainframe');
let startCallButton = document.getElementById('start-call-button');
let hangUpCallButton = document.getElementById('hangup-call-button');
let acceptCallButton = document.getElementById('accept-call-button');
let startVideoButton = document.getElementById('start-video-button');
let stopVideoButton = document.getElementById('stop-video-button');
let connectedLabel = document.getElementById('connectedLabel');
let remoteVideoContainer = document.getElementById('remoteVideoContainer');
let localVideoContainer = document.getElementById('localVideoContainer');
/**
 * Create an instance of CallClient. Initialize a TeamsCallAgent instance with a CommunicationUserCredential via created CallClient. TeamsCallAgent enables us to make outgoing calls and receive incoming calls. 
 * You can then use the CallClient.getDeviceManager() API instance to get the DeviceManager.
 */
initializeCallAgentButton.onclick = async () => {
    try {
        const callClient = new CallClient(); 
        //tokenCredential = new AzureCommunicationTokenCredential(userAccessToken.value.trim());
        tokenCredential = new AzureCommunicationTokenCredential('eyJhbGciOiJSUzI1NiIsImtpZCI6IjExRkNCRjhEQzBFRTMzQUY3QkIwQTE3OUUzNjI0RUNBNjk1ODE2NjQiLCJ4NXQiOiJFZnlfamNEdU02OTdzS0Y1NDJKT3ltbFlGbVEiLCJ0eXAiOiJKV1QifQ.eyJza3lwZWlkIjoib3JnaWQ6OGE5YjRiMmYtNjA0MC00ZTJiLWIyMjktZTZmOTEzM2VmZWQyIiwic2NwIjoxMDI0LCJjc2kiOiIxNzMzOTk4NTQ3IiwiZXhwIjoxNzM0MDAzMzg2LCJyZ24iOiJhbWVyIiwidGlkIjoiZmZkMGE5MDctMGEyZi00NGIwLTg1ZjctMmY2M2M3YzA3MzMzIiwiYWNzU2NvcGUiOiJ2b2lwLGNoYXQiLCJyZXNvdXJjZUlkIjoiYmViZjcyMjAtMDEzYS00MzEzLWE5MDEtNDA2MTZiMTIyOTYwIiwiYWFkX2lhdCI6IjE3MzM5OTg1NDciLCJhYWRfdXRpIjoiZXVxV1BrbEZDVS0xWVlocmwwd3FBQSIsImFhZF9hcHBpZCI6IjFmZDUxMThlLTI1NzYtNDI2My04MTMwLTk1MDMwNjRjODM3YSIsImlhdCI6MTczMzk5ODg0OCwiYWFkX3BmdCI6IjJxMVp5VkxKU01sVFNVWElDMGg3QmdTRlZYdVpKLWU2aC1ha3VFUWJoRWFXLVhqNFpqdG5Cb1lHT3hoSGhodm1HNXNGVllWa3V2Z1pBSGM1UW5TNUF1cW9pTmQwbzNOLXZJQ1M3M01fVVBUYzFNcmswSktYRTJjd0xwTktWZ0lwYUFBfGhWTnJiOW93RlAwclZiNE9neDA3ZmlCTldsSkt5eWkwUUhtMDA0UVNiTk9rTkJBQ2hURHR2LTg2bGZwMWtlVmNINTlqWDU5cl9fRWlyLTNGMUd0NDF4Q2taVWxhMW1vY0t5d1FqbjJMR0Vzd2tvRVZ5TGVjcnNRS0MwcHBDd1EzWGhzM3ZLN1hKb0FvSlFNbUd0NXRQV1FZVXlwNXc4dHJUZy1XSnFDNGczODRHNFZGU3o2SDhEMjhUZ3Bfc0RaZGNYcWFNN0VwYlBwNEh0a3FIVTgydzBsMmF0MHRkbGxfTUJzZlNrUHMtT1d5ZVM3VzV1T1lUYXZVdkViRnBYd3MtdjNkN2U3bGdIWExyUE5PcHFPYjhOdGdxTGw0eU1lVnJiNkpwOVpnMnBsLWgtMzdYdnRYblFmM2ZqZTgtenExSWFUa0E3VDlPb1prQUdjQV80ek5tOW5BM0ZzOXlDRXFJVHFXWmctaGhsRGlKcUc4NmJzRzBPaVRkX1VsbkR0T3JCS1dnSkVjTTR5WThST1UtTDVDaGx0RktEWFdhTGZfdTdNSVhQTXg5Z01xV0FlVGE0QmpCemZEQlE0am5vLVdsZXlYTjVPdXFsUjJsdEZsUUt6cXJzN2xmYVNqb2xkMVNqM2JUd1FuWVUtck1HeUNmdXdjUHg1ZXRfdjBFaF9TYmQ0Y20xalA5LW5CWExtRExOUDNuZG1YMjd5ZUJNRVJCR0sxU3FTR2xBMkxJV1hpY3lRWklTaXdBVlRWUUFzWVVHZEFYWmQyLUVEelludTVMQmJzbkJVanZGMFNzN0R6S004MnQxM1QyLWZYdGw5WmprSHhBWXJwQklKRjdmMHpETE1zQnE5LVpMVmxxTnh1amk2UnNxa04wRGJfSTd6VTZ5UzFTVzZESjFmaEUzT0ducWgwdlhUd1NZZ2EtZXpkRFRoSkZfdUNjU3dsUlVSYmdoaFhCQ1ZRQWtRVG9UaWpKakRjdll5VGNFc0JaaFBMTklLYUtTQkxoU1JoRkFtZUVNV01EQUtzM0wwNnY1ZkxWT19oQWtCV1Y5TDctdzh8YU9EOXN1OFhiV3g2NVVpMnVjN1V3ZXZMd3Q2MDI4SHN6U3pFQ0VnaWE0VXBlajBobUtsN3JBUm5NaDNfLXRFbl9NUjFGcFdKZmRYQTlXMThPSzRhcVRpVVpzOGFFa201OVc1dHJFZFEzclBZREhVeEhzNEhnU2tmWWlIajQxMlhlb3NPU05TcEV6dXV0dkJQbWZLdF9MSmViX2tjWXN3QjdiNDJDcFpQVGxaSENLMnoyNHZEeFJWdE9JODhvbWl1aUx2MVprSC1sQUFaWHlEZTlDc2pMZkl3Rk14eXJjeWphRHdvZTdtckU0YWw5NE5PWW1xd3VzYlBvWTRLSl9rRWE0N1JuaF9fN2dIWnNfTjVCN1NIUDc0MzhkLWwzQnNCbGczZWplOTRqV3l0dFB2aDFjVURjcnlNMzl6R0lCcGszNzVvWkQ3Y0VVVE5kOENQWk0xQWxnIn0.mvM44dEDhSlMiTJdhoPWidC1H1QCYL3cO_CcgGNR4DfUxeKCaM6ozBgQSPUxGZ82Qtu9stGZ7SgOUnOYAYOMkjX3k8DO_fG2ZMXOV5PxTl0KJAd-KzY7CsAs-MBzwIomBey4cAPaZBTjolkEGMxhaR5DIe-8NeeCW14n5kfkTt-pUtk2UX88I__D3lkjZAYAOvjPmTnIJu0wYgWicSUykkUSY0Im6Y6Gc6j5Hl1LuqG1H5QGB36eb_1hmeG8Cf1pDwMHnevL2BmHgibJouCibjPgWzkFXseSmKRfyRVABjULo6mT0s1fVgZ7E7RDAg4aD2OpoeRzZNX8GSNFvxODsA'.trim());
        teamsCallAgent = await callClient.createTeamsCallAgent(tokenCredential)
        // Set up a camera device to use.
        deviceManager = await callClient.getDeviceManager();
        await deviceManager.askDevicePermission({ video: true });
        await deviceManager.askDevicePermission({ audio: true });
        // Listen for an incoming call to accept.
        teamsCallAgent.on('incomingCall', async (args) => {
            try {
                incomingCall = args.incomingCall;
                acceptCallButton.disabled = false;
                startCallButton.disabled = false;
            } catch (error) {
                console.error(error);
            }
        });
        startCallButton.disabled = false;
        initializeCallAgentButton.disabled = true;
        window.alert("Init successful");
    } catch(error) {
        console.error(error);
        window.alert("Init not successful");
    }
}
/**
 * Place a 1:1 outgoing video call to a user
 * Add an event listener to initiate a call when the `startCallButton` is selected.
 * Enumerate local cameras using the deviceManager `getCameraList` API.
 * In this quickstart, we're using the first camera in the collection. Once the desired camera is selected, a
 * LocalVideoStream instance will be constructed and passed within `videoOptions` as an item within the
 * localVideoStream array to the call method. When the call connects, your application will be sending a video stream to the other participant. 
 */
startCallButton.onclick = async () => {
    try {
        const localVideoStream = await createLocalVideoStream();
        const videoOptions = localVideoStream ? { localVideoStreams: [localVideoStream] } : undefined;
        //call = teamsCallAgent.startCall({ microsoftTeamsUserId: calleeTeamsUserId.value.trim() }, { videoOptions: videoOptions });
        //call = teamsCallAgent.startCall({ microsoftTeamsUserId: '4d52edce-fc9a-4c32-836d-e63ada743ad0'.trim() }, { videoOptions: videoOptions });
        // Marcos ID
        //call = teamsCallAgent.startCall({ microsoftTeamsUserId: '69981ef7-d5d7-4f66-bc7f-e4d0ded559a7'.trim() }, { videoOptions: videoOptions });
        //Jans ID
        call = teamsCallAgent.startCall({ microsoftTeamsUserId: 'da9054b5-28b8-41c4-a887-2fa5b8bcfeb6'.trim() }, { videoOptions: videoOptions });
        // Subscribe to the call's properties and events.
        subscribeToCall(call);
        mainframe.hidden = true;
    } catch (error) {
        console.error(error);
        window.alert("Call not successful");
    }
}
/**
 * Accepting an incoming call with a video
 * Add an event listener to accept a call when the `acceptCallButton` is selected.
 * You can accept incoming calls after subscribing to the `TeamsCallAgent.on('incomingCall')` event.
 * You can pass the local video stream to accept the call with the following code.
 */
acceptCallButton.onclick = async () => {
    try {
        const localVideoStream = await createLocalVideoStream();
        const videoOptions = localVideoStream ? { localVideoStreams: [localVideoStream] } : undefined;
        call = await incomingCall.accept({ videoOptions });
        // Subscribe to the call's properties and events.
        subscribeToCall(call);
    } catch (error) {
        console.error(error);
    }
}
// Subscribe to a call obj.
// Listen for property changes and collection udpates.
subscribeToCall = (call) => {
    try {
        // Inspect the initial call.id value.
        console.log(`Call Id: ${call.id}`);
        //Subsribe to call's 'idChanged' event for value changes.
        call.on('idChanged', () => {
            console.log(`Call ID changed: ${call.id}`); 
        });
        // Inspect the initial call.state value.
        console.log(`Call state: ${call.state}`);
        // Subscribe to call's 'stateChanged' event for value changes.
        call.on('stateChanged', async () => {
            console.log(`Call state changed: ${call.state}`);
            if(call.state === 'Connected') {
                connectedLabel.hidden = false;
                acceptCallButton.disabled = true;
                startCallButton.disabled = true;
                hangUpCallButton.disabled = false;
                startVideoButton.disabled = false;
                stopVideoButton.disabled = false;
            } else if (call.state === 'Disconnected') {
                connectedLabel.hidden = true;
                startCallButton.disabled = false;
                hangUpCallButton.disabled = true;
                startVideoButton.disabled = true;
                stopVideoButton.disabled = true;
                console.log(`Call ended, call end reason={code=${call.callEndReason.code}, subCode=${call.callEndReason.subCode}}`);
            }   
        });
        call.on('isLocalVideoStartedChanged', () => {
            console.log(`isLocalVideoStarted changed: ${call.isLocalVideoStarted}`);
        });
        console.log(`isLocalVideoStarted: ${call.isLocalVideoStarted}`);
        call.localVideoStreams.forEach(async (lvs) => {
            localVideoStream = lvs;
            await displayLocalVideoStream();
        });
        call.on('localVideoStreamsUpdated', e => {
            e.added.forEach(async (lvs) => {
                localVideoStream = lvs;
                await displayLocalVideoStream();
            });
            e.removed.forEach(lvs => {
               removeLocalVideoStream();
            });
        });
        
        // Inspect the call's current remote participants and subscribe to them.
        call.remoteParticipants.forEach(remoteParticipant => {
            subscribeToRemoteParticipant(remoteParticipant);
        });
        // Subscribe to the call's 'remoteParticipantsUpdated' event to be
        // notified when new participants are added to the call or removed from the call.
        call.on('remoteParticipantsUpdated', e => {
            // Subscribe to new remote participants that are added to the call.
            e.added.forEach(remoteParticipant => {
                subscribeToRemoteParticipant(remoteParticipant)
            });
            // Unsubscribe from participants that are removed from the call
            e.removed.forEach(remoteParticipant => {
                console.log('Remote participant removed from the call.');
            });
        });
    } catch (error) {
        console.error(error);
    }
}
// Subscribe to a remote participant obj.
// Listen for property changes and collection udpates.
subscribeToRemoteParticipant = (remoteParticipant) => {
    try {
        // Inspect the initial remoteParticipant.state value.
        console.log(`Remote participant state: ${remoteParticipant.state}`);
        // Subscribe to remoteParticipant's 'stateChanged' event for value changes.
        remoteParticipant.on('stateChanged', () => {
            console.log(`Remote participant state changed: ${remoteParticipant.state}`);
        });
        // Inspect the remoteParticipants's current videoStreams and subscribe to them.
        remoteParticipant.videoStreams.forEach(remoteVideoStream => {
            subscribeToRemoteVideoStream(remoteVideoStream)
        });
        // Subscribe to the remoteParticipant's 'videoStreamsUpdated' event to be
        // notified when the remoteParticiapant adds new videoStreams and removes video streams.
        remoteParticipant.on('videoStreamsUpdated', e => {
            // Subscribe to newly added remote participant's video streams.
            e.added.forEach(remoteVideoStream => {
                subscribeToRemoteVideoStream(remoteVideoStream)
            });
            // Unsubscribe from newly removed remote participants' video streams.
            e.removed.forEach(remoteVideoStream => {
                console.log('Remote participant video stream was removed.');
            })
        });
    } catch (error) {
        console.error(error);
    }
}
/**
 * Subscribe to a remote participant's remote video stream obj.
 * You have to subscribe to the 'isAvailableChanged' event to render the remoteVideoStream. If the 'isAvailable' property
 * changes to 'true' a remote participant is sending a stream. Whenever the availability of a remote stream changes
 * you can choose to destroy the whole 'Renderer' a specific 'RendererView' or keep them. Displaying RendererView without a video stream will result in a blank video frame. 
 */
subscribeToRemoteVideoStream = async (remoteVideoStream) => {
    // Create a video stream renderer for the remote video stream.
    let videoStreamRenderer = new VideoStreamRenderer(remoteVideoStream);
    let view;
    console.log("Start Remote Video");
    const renderVideo = async () => {
        try {
            // Create a renderer view for the remote video stream.
            view = await videoStreamRenderer.createView();
            // Attach the renderer view to the UI.
            remoteVideoContainer.hidden = false;
            mainframe.hidden = true;
            remoteVideoContainer.appendChild(view.target);
        } catch (e) {
            console.warn(`Failed to createView, reason=${e.message}, code=${e.code}`);
        }	
    }
    
    remoteVideoStream.on('isAvailableChanged', async () => {
        // Participant has switched video on.
        if (remoteVideoStream.isAvailable) {
            await renderVideo();
        // Participant has switched video off.
        } else {
            if (view) {
                //view.dispose();

                mainframe.hidden = false;
                view = undefined;
            }
        }
    });
    // Participant has video on initially.
    if (remoteVideoStream.isAvailable) {
        await renderVideo();
    }
}
// Start your local video stream.
// This will send your local video stream to remote participants so they can view it.
startVideoButton.onclick = async () => {
    try {
        const localVideoStream = await createLocalVideoStream();
        await call.startVideo(localVideoStream);
    } catch (error) {
        console.error(error);
    }
}
// Stop your local video stream.
// This will stop your local video stream from being sent to remote participants.
stopVideoButton.onclick = async () => {
    try {
        await call.stopVideo(localVideoStream);
    } catch (error) {
        console.error(error);
    }
}
/**
 * To render a LocalVideoStream, you need to create a new instance of VideoStreamRenderer, and then
 * create a new VideoStreamRendererView instance using the asynchronous createView() method.
 * You may then attach view.target to any UI element. 
 */
// Create a local video stream for your camera device
createLocalVideoStream = async () => {

    console.log("deviceManager");
    console.log(deviceManager);
    const camera = (await deviceManager.getCameras())[0];
    if (camera) {
        return new LocalVideoStream(camera);
    } else {
        console.error(`No camera device found on the system`);
    }
}
// Display your local video stream preview in your UI
displayLocalVideoStream = async () => {
    try {
        localVideoStreamRenderer = new VideoStreamRenderer(localVideoStream);
        const view = await localVideoStreamRenderer.createView();
        localVideoContainer.hidden = false;
        localVideoContainer.appendChild(view.target);
    } catch (error) {
        console.error(error);
    } 
}
// Remove your local video stream preview from your UI
removeLocalVideoStream = async() => {
    try {
        localVideoStreamRenderer.dispose();
        localVideoContainer.hidden = true;
    } catch (error) {
        console.error(error);
    } 
}
// End the current call
hangUpCallButton.addEventListener("click", async () => {
    // end the current call
    await call.hangUp();
});

/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. */

import { CustomError } from './custom.error';

export class Utilities {
    static log(exception: Error | CustomError | string) {
        if (exception == null) {
    
          console.error(exception);
    
        }
    
        else if (typeof exception === 'string') {
    
          console.error(exception);
    
        }
    
        else {
    
          console.group(`${exception.name}: ${exception.message}`);
    
          {
    
            let innerException = exception;
    
            if (exception instanceof CustomError) {
    
              innerException = exception.innerError;
    
            }
    
            if ((window as any).OfficeExtension && innerException instanceof OfficeExtension.Error) {
    
              console.groupCollapsed('Debug Info');
    
              console.error(innerException.debugInfo);
    
              console.groupEnd();
    
            }
    
            {
    
              console.groupCollapsed('Stack Trace');
    
              console.error(exception.stack);
    
              console.groupEnd();
    
            }
    
            {
    
              console.groupCollapsed('Inner Error');
    
              console.error(innerException);
    
              console.groupEnd();
    
            }
    
          }
    
          console.groupEnd();
    
        }
    
    }

    static notify(error: Error) {
        if (error == null) {
    
          console.error(new Error('Invalid params. Cannot create a notification'));
    
          return null;
    
        }

        const existingNotifications = document.getElementsByClassName('helpers-notification');
    
        while (existingNotifications[0]) {
    
          existingNotifications[0].parentNode.removeChild(existingNotifications[0]);
    
        }

        const html:string = "<div class=\"helpers-notification ms-MessageBar ms-MessageBar--error\"><button><i class=\"ms-Icon ms-Icon--Clear\"></i></button></div>";
        document.body.insertAdjacentHTML("afterbegin", html);
  
        const notificationDiv = document.getElementsByClassName('helpers-notification')[0];
    
        const messageTextArea = document.createElement('div');
    
        notificationDiv.insertAdjacentElement('beforeend', messageTextArea);
    
    
    
        if (error.name) {
    
          const titleDiv = document.createElement('div');
    
          titleDiv.textContent = error.name;
    
          titleDiv.classList.add('ms-fontWeight-semibold');
    
          messageTextArea.insertAdjacentElement('beforeend', titleDiv);
    
        }
    
    
        error.message.split('\n').forEach(text => {
    
          const div = document.createElement('div');
    
          div.textContent = text;
    
          messageTextArea.insertAdjacentElement('beforeend', div);
    
        });
    
    
    
        if (error.stack) {
    
          const labelDiv = document.createElement('div');
    
          messageTextArea.insertAdjacentElement('beforeend', labelDiv);
    
          const label = document.createElement('a');
    
          label.setAttribute('href', 'javascript:void(0)');
    
          label.onclick = () => {
    
            (document.querySelector('.helpers-notification pre') as HTMLPreElement)
    
              .parentElement.style.display = 'block';
    
            labelDiv.style.display = 'none';
    
          };
    
          label.textContent = 'Details';
    
          labelDiv.insertAdjacentElement('beforeend', label);
    
    
    
          const preDiv = document.createElement('div');
    
          preDiv.style.display = 'none';
    
          messageTextArea.insertAdjacentElement('beforeend', preDiv);
    
          const detailsDiv = document.createElement('pre');
    
          detailsDiv.textContent = error.stack;
    
          preDiv.insertAdjacentElement('beforeend', detailsDiv);
    
        }
    
    
    
        (document.querySelector('.helpers-notification > button') as HTMLButtonElement)
    
          .onclick = () => notificationDiv.parentNode.removeChild(notificationDiv);
    
      }
    

}
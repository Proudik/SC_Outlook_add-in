/* global Office */
import { onMessageSendHandler } from "./onMessageSendHandler";

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
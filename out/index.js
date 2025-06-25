"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MicrosoftGraph = void 0;
var identity_1 = require("@azure/identity");
var microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
var azureTokenCredentials_1 = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
var MicrosoftGraph = /** @class */ (function () {
    function MicrosoftGraph() {
    }
    MicrosoftGraph.reloadInstanceForGroup = function (groupId) {
        delete MicrosoftGraph._instances[groupId];
    };
    MicrosoftGraph.getUserId = function (groupId, userEmail) {
        return __awaiter(this, void 0, void 0, function () {
            var graph, userId;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, MicrosoftGraph._getInstance(groupId, true)];
                    case 1:
                        graph = _a.sent();
                        if (!graph) return [3 /*break*/, 3];
                        return [4 /*yield*/, graph._getUserIdFromEmail(userEmail)];
                    case 2:
                        userId = _a.sent();
                        return [2 /*return*/, userId];
                    case 3: return [2 /*return*/, null];
                }
            });
        });
    };
    MicrosoftGraph.isUserAuthorizedForUpSignOn = function (groupId, userId) {
        return __awaiter(this, void 0, void 0, function () {
            var graph, isAuthorized;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, MicrosoftGraph._getInstance(groupId, false)];
                    case 1:
                        graph = _a.sent();
                        if (!graph) return [3 /*break*/, 3];
                        return [4 /*yield*/, graph.isUserAuthorizedForUpSignOn(userId)];
                    case 2:
                        isAuthorized = _a.sent();
                        return [2 /*return*/, isAuthorized];
                    case 3: return [2 /*return*/, false];
                }
            });
        });
    };
    MicrosoftGraph.getGroupsForUser = function (groupId, userId) {
        return __awaiter(this, void 0, void 0, function () {
            var graph, groups;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, MicrosoftGraph._getInstance(groupId, false)];
                    case 1:
                        graph = _a.sent();
                        if (!graph) return [3 /*break*/, 3];
                        return [4 /*yield*/, graph.getGroupsForUser(userId)];
                    case 2:
                        groups = _a.sent();
                        return [2 /*return*/, groups];
                    case 3: return [2 /*return*/, []];
                }
            });
        });
    };
    MicrosoftGraph.getAllUsersAssignedToUpSignOn = function (groupId, withoutConfigRefresh) {
        return __awaiter(this, void 0, void 0, function () {
            var graph, allUsers;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, MicrosoftGraph._getInstance(groupId, withoutConfigRefresh)];
                    case 1:
                        graph = _a.sent();
                        if (!graph) return [3 /*break*/, 3];
                        return [4 /*yield*/, graph.getAllUsersAssignedToUpSignOn()];
                    case 2:
                        allUsers = _a.sent();
                        return [2 /*return*/, allUsers];
                    case 3: return [2 /*return*/, []];
                }
            });
        });
    };
    MicrosoftGraph._getInstance = function (groupId, withoutConfigRefresh) {
        return __awaiter(this, void 0, void 0, function () {
            var entraConfig;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!withoutConfigRefresh && MicrosoftGraph._instances[groupId]) {
                            return [2 /*return*/, MicrosoftGraph._instances[groupId]];
                        }
                        return [4 /*yield*/, MicrosoftGraph.getMSEntraConfigForGroup(groupId)];
                    case 1:
                        entraConfig = _a.sent();
                        if (!MicrosoftGraph._instances[groupId] || MicrosoftGraph._hasConfigChanged(groupId, entraConfig)) {
                            if ((entraConfig === null || entraConfig === void 0 ? void 0 : entraConfig.tenantId) && entraConfig.clientId && entraConfig.clientSecret && entraConfig.appResourceId) {
                                MicrosoftGraph._instances[groupId] = new _MicrosoftGraph(entraConfig.tenantId, entraConfig.clientId, entraConfig.clientSecret, entraConfig.appResourceId);
                            }
                            else {
                                delete MicrosoftGraph._instances[groupId];
                            }
                            MicrosoftGraph._groupConfig[groupId] = entraConfig;
                        }
                        return [2 /*return*/, MicrosoftGraph._instances[groupId] || null];
                }
            });
        });
    };
    MicrosoftGraph._hasConfigChanged = function (groupId, currentConfig) {
        var cachedConfig = MicrosoftGraph._groupConfig[groupId];
        if (currentConfig == null && cachedConfig == null)
            return false;
        if ((currentConfig === null || currentConfig === void 0 ? void 0 : currentConfig.tenantId) != (cachedConfig === null || cachedConfig === void 0 ? void 0 : cachedConfig.tenantId) ||
            (currentConfig === null || currentConfig === void 0 ? void 0 : currentConfig.clientId) != (cachedConfig === null || cachedConfig === void 0 ? void 0 : cachedConfig.clientId) ||
            (currentConfig === null || currentConfig === void 0 ? void 0 : currentConfig.clientSecret) != (cachedConfig === null || cachedConfig === void 0 ? void 0 : cachedConfig.clientSecret) ||
            (currentConfig === null || currentConfig === void 0 ? void 0 : currentConfig.appResourceId) != (cachedConfig === null || cachedConfig === void 0 ? void 0 : cachedConfig.appResourceId)) {
            return true;
        }
        return false;
    };
    MicrosoftGraph.listNeededAPIs = function () {
        return [
            {
                path: "/users",
                docLink: "https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/groups",
                docLink: "https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/users/{id | userPrincipalName}/appRoleAssignments",
                docLink: "https://learn.microsoft.com/en-us/graph/api/user-list-approleassignments?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/servicePrincipals/{id}/appRoleAssignedTo",
                docLink: "https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignedto?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/groups/{id}/members/microsoft.graph.user",
                docLink: "https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0&tabs=http",
            },
            {
                path: "/users/{id}/memberOf/microsoft.graph.group",
                docLink: "https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0&tabs=http",
            },
        ];
    };
    MicrosoftGraph._instances = {};
    MicrosoftGraph._groupConfig = {};
    return MicrosoftGraph;
}());
exports.MicrosoftGraph = MicrosoftGraph;
var _MicrosoftGraph = /** @class */ (function () {
    /**
     *
     * @param tenantId - The Microsoft Entra tenant (directory) ID.
     * @param clientId - The client (application) ID of an App Registration in the tenant.
     * @param clientSecret - A client secret that was generated for the App Registration.
     * @param appResourceId - Identifier of the ressource (UpSignOn) in the graph that users need to have access to in order to be authorized to use an UpSignOn licence
     */
    function _MicrosoftGraph(tenantId, clientId, clientSecret, appResourceId) {
        var credential = new identity_1.ClientSecretCredential(tenantId, clientId, clientSecret);
        var authProvider = new azureTokenCredentials_1.TokenCredentialAuthenticationProvider(credential, {
            // The client credentials flow requires that you request the
            // /.default scope, and pre-configure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            scopes: ["https://graph.microsoft.com/.default"],
        });
        var clientOptions = {
            authProvider: authProvider,
        };
        this.msGraph = microsoft_graph_client_1.Client.initWithMiddleware(clientOptions);
        this.appResourceId = appResourceId;
    }
    /**
     * Gets the id of the first user to match that email address and who has been assigned the role for using UpSignOn
     *
     * @param email
     * @returns the id if such a user exists, null otherwise
     */
    _MicrosoftGraph.prototype._getUserIdFromEmail = function (email) {
        return __awaiter(this, void 0, void 0, function () {
            var users, userId;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!email.match(/^[\w-\.+]+@([\w-]+\.)+[\w-]{2,4}$/)) {
                            throw "Email is malformed";
                        }
                        return [4 /*yield*/, this.msGraph
                                // PERMISSION = User.Read.All OR Directory.Read.All
                                .api("/users")
                                .header("ConsistencyLevel", "eventual")
                                .filter("mail eq '".concat(email, "' or userPrincipalName eq '").concat(email, "' or otherMails/any(oe:oe eq '").concat(email, "')"))
                                .select(["id"])
                                .get()];
                    case 1:
                        users = _b.sent();
                        userId = (_a = users.value[0]) === null || _a === void 0 ? void 0 : _a.id;
                        return [2 /*return*/, userId];
                }
            });
        });
    };
    _MicrosoftGraph.prototype.isUserAuthorizedForUpSignOn = function (userId) {
        return __awaiter(this, void 0, void 0, function () {
            var allAuthorizedUserIds;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.getAllUsersAssignedToUpSignOn()];
                    case 1:
                        allAuthorizedUserIds = _a.sent();
                        return [2 /*return*/, allAuthorizedUserIds.indexOf(userId) >= 0];
                }
            });
        });
    };
    _MicrosoftGraph.prototype.getAllUsersAssignedToUpSignOn = function () {
        return __awaiter(this, void 0, void 0, function () {
            var allPrincipalsRes, allUsersId, allGroups, i, g, allGroupUsersRes;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.msGraph
                            // PERMISSION = Application.Read.All OR Directory.Read.All
                            // https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignedto?view=graph-rest-1.0&tabs=http
                            .api("/servicePrincipals/".concat(this.appResourceId, "/appRoleAssignedTo"))
                            .header("ConsistencyLevel", "eventual")
                            .select(["principalType", "principalId"])
                            .get()];
                    case 1:
                        allPrincipalsRes = _a.sent();
                        allUsersId = allPrincipalsRes.value
                            .filter(function (u) { return u.principalType === "User"; })
                            .map(function (u) { return u.principalId; });
                        allGroups = allPrincipalsRes.value.filter(function (u) { return u.principalType === "Group"; });
                        i = 0;
                        _a.label = 2;
                    case 2:
                        if (!(i < allGroups.length)) return [3 /*break*/, 5];
                        g = allGroups[i];
                        return [4 /*yield*/, this.listGroupMembers(g.principalId)];
                    case 3:
                        allGroupUsersRes = _a.sent();
                        allUsersId = __spreadArray(__spreadArray([], allUsersId, true), allGroupUsersRes.map(function (u) { return u.id; }), true);
                        _a.label = 4;
                    case 4:
                        i++;
                        return [3 /*break*/, 2];
                    case 5: return [2 /*return*/, allUsersId];
                }
            });
        });
    };
    /**
     * Returns all groups (and associated groups) that this user belongs to
     * To be used for sharing to teams ?
     * This would suppose a user can only shared to teams to which it belongs ?
     * @param email
     * @returns
     */
    _MicrosoftGraph.prototype.getGroupsForUser = function (userId) {
        return __awaiter(this, void 0, void 0, function () {
            var groups;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.msGraph
                            // Get groups, directory roles, and administrative units that the user is a transitive member of.
                            // PERMISSION = Directory.Read.All OR GroupMember.Read.All OR Directory.Read.All
                            // https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0&tabs=http
                            // .api(`/users/${userId}/memberOf`) // pour tout avoir
                            // .api(`/users/${userId}/memberOf/microsoft.graph.administrativeUnit`) // pour avoir tous les administrativeUnit
                            .api("/users/".concat(userId, "/transitiveMemberOf/microsoft.graph.group")) // pour avoir tous les groupes
                            .header("ConsistencyLevel", "eventual")
                            .select(["id", "displayName"])
                            .get()];
                    case 1:
                        groups = _a.sent();
                        return [2 /*return*/, groups.value];
                }
            });
        });
    };
    /**
     * Returns all members of a group
     * @returns
     */
    _MicrosoftGraph.prototype.listGroupMembers = function (groupId) {
        return __awaiter(this, void 0, void 0, function () {
            var groupMembers;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.msGraph
                            .api("/groups/".concat(groupId, "/transitiveMembers/microsoft.graph.user/"))
                            .header("ConsistencyLevel", "eventual")
                            .select(["id", "mail", "displayName"])
                            .get()];
                    case 1:
                        groupMembers = _a.sent();
                        return [2 /*return*/, groupMembers.value];
                }
            });
        });
    };
    _MicrosoftGraph.prototype.checkGroupMembers = function (groupIds) {
        return __awaiter(this, void 0, void 0, function () {
            var allGroups;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.msGraph
                            .api("/groups")
                            .header("ConsistencyLevel", "eventual")
                            .filter("id in ('".concat(groupIds.join("', '"), "')"))
                            .expand("members($select=id, displayName, mail)")
                            .select(["id", "displayName"])
                            .get()];
                    case 1:
                        allGroups = _a.sent();
                        // beware, that mail could be empty although the user may have another email
                        return [2 /*return*/, allGroups.value];
                }
            });
        });
    };
    return _MicrosoftGraph;
}());

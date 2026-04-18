export namespace main {
	
	export class AppState {
	    processed_customers: Record<string, number>;
	    target_userid: string;
	    fixed_members: string[];
	    group_owner: string;
	    need_review_list?: string[];
	    test_customer_names?: string[];
	
	    static createFrom(source: any = {}) {
	        return new AppState(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.processed_customers = source["processed_customers"];
	        this.target_userid = source["target_userid"];
	        this.fixed_members = source["fixed_members"];
	        this.group_owner = source["group_owner"];
	        this.need_review_list = source["need_review_list"];
	        this.test_customer_names = source["test_customer_names"];
	    }
	}
	export class Contact {
	    external_userid: string;
	    name: string;
	    type: number;
	    corp_name: string;
	
	    static createFrom(source: any = {}) {
	        return new Contact(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.external_userid = source["external_userid"];
	        this.name = source["name"];
	        this.type = source["type"];
	        this.corp_name = source["corp_name"];
	    }
	}
	export class GroupChat {
	    chat_id: string;
	    name: string;
	    owner: string;
	    member_count: number;
	
	    static createFrom(source: any = {}) {
	        return new GroupChat(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.chat_id = source["chat_id"];
	        this.name = source["name"];
	        this.owner = source["owner"];
	        this.member_count = source["member_count"];
	    }
	}
	export class GroupResult {
	    Success: boolean;
	    PrivacySet: boolean;
	    PrivacyVerified: boolean;
	    MembersSelected: number;
	    MembersExpected: number;
	    ErrorDetail: string;
	    NeedManualCheck: boolean;
	
	    static createFrom(source: any = {}) {
	        return new GroupResult(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.Success = source["Success"];
	        this.PrivacySet = source["PrivacySet"];
	        this.PrivacyVerified = source["PrivacyVerified"];
	        this.MembersSelected = source["MembersSelected"];
	        this.MembersExpected = source["MembersExpected"];
	        this.ErrorDetail = source["ErrorDetail"];
	        this.NeedManualCheck = source["NeedManualCheck"];
	    }
	}
	export class Member {
	    userid: string;
	    name: string;
	    status: number;
	
	    static createFrom(source: any = {}) {
	        return new Member(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.userid = source["userid"];
	        this.name = source["name"];
	        this.status = source["status"];
	    }
	}

}

